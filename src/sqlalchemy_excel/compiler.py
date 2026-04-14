"""SQL compiler for the Excel dialect.

Compiles SQLAlchemy expression trees into the SQL subset that
excel-dbapi's parser understands:

Supported:
    SELECT columns FROM table [WHERE ...] [GROUP BY ...] [HAVING ...] [ORDER BY col [ASC|DESC]] [LIMIT n] [OFFSET n]
    SELECT DISTINCT columns FROM table [WHERE ...] ...
    SELECT cols FROM t1 [INNER|LEFT|RIGHT|FULL OUTER] JOIN t2 ON ... { JOIN t3 ON ... } [WHERE ...] [ORDER BY ...] [LIMIT n] [OFFSET n]
    SELECT cols FROM t1 CROSS JOIN t2 [WHERE ...] [ORDER BY ...] [LIMIT n] [OFFSET n]
    UNION / UNION ALL / INTERSECT / EXCEPT compound SELECT
    INSERT INTO table (cols) VALUES (vals), (vals2), ...
    INSERT INTO table (cols) SELECT cols FROM source [WHERE ...]
    UPDATE table SET col=val [WHERE ...]
    DELETE FROM table [WHERE ...]

    Rejected (raises CompileError):
    CTEs, window functions, aggregate FILTER, RETURNING, FOR UPDATE,
    NATURAL JOIN, non-equality/OR/non-column ON clauses,
    DISTINCT with JOIN, grouped/nested compound branches

    Partially supported:
    non-correlated subqueries in WHERE ... IN (SELECT single_col FROM table [WHERE ...])
    — supported in SELECT, UPDATE, and DELETE

For single-table queries, excel-dbapi expects unquoted, unprefixed column names:
    SELECT id, name FROM users          (correct)
    SELECT users.id, users.name FROM users  (WRONG — parser rejects)

For JOIN queries, table-qualified column names are required:
    SELECT users.id, orders.user_id FROM users JOIN orders ON users.id = orders.user_id

The compiler detects JOIN context and switches between the two modes automatically.
"""

from __future__ import annotations

import itertools
import re
from typing import TYPE_CHECKING, Any, Literal, cast

from sqlalchemy import exc
from sqlalchemy.sql import (
    coercions,
    compiler,
    elements,
    operators,
    roles,
    visitors,
)
from sqlalchemy.sql.expression import Join

if TYPE_CHECKING:
    from collections.abc import Callable, MutableMapping


_SUPPORTED_FUNCTIONS = {"count", "sum", "avg", "min", "max"}
_BARE_IDENTIFIER_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*\Z")


class ExcelIdentifierPreparer(compiler.IdentifierPreparer):
    """Identifier preparer that never quotes identifiers.

    excel-dbapi's parser expects bare column names without quotes
    or table prefixes.
    """

    def __init__(self, dialect: Any) -> None:
        super().__init__(dialect, initial_quote="", final_quote="")
        self.reserved_words = set()

    @staticmethod
    def _validate_identifier(ident: str) -> None:
        if not _BARE_IDENTIFIER_RE.fullmatch(ident):
            raise exc.CompileError(
                "Excel dialect identifiers must match [A-Za-z_][A-Za-z0-9_]*"
            )

    def quote_identifier(self, value: str) -> str:
        self._validate_identifier(value)
        return value

    def quote(self, ident: str, force: Any = None) -> str:
        self._validate_identifier(ident)
        return ident


class ExcelCompiler(compiler.SQLCompiler):
    """Compiles SQLAlchemy SQL expressions for excel-dbapi."""

    _in_in_clause: bool = False
    _subquery_depth: int = 0
    _has_join: bool = False

    @staticmethod
    def _raise_if_schema_qualified(table: Any) -> None:
        """Reject schema-qualified SQLAlchemy table objects.

        This guard runs only for SQL emitted from SQLAlchemy expression
        compilation. Raw SQL strings executed through ``exec_driver_sql()`` are
        passed directly to excel-dbapi and intentionally bypass SQLAlchemy
        compile-time schema validation.
        """
        if getattr(table, "schema", None) is not None:
            raise exc.CompileError("Excel dialect does not support schemas")

    def visit_table(
        self,
        table: Any,
        asfrom: bool = False,
        iscrud: bool = False,
        ashint: bool = False,
        fromhints: Any = None,
        use_schema: bool = True,
        from_linter: Any = None,
        ambiguous_table_name_map: Any = None,
        enclosing_alias: Any = None,
        **kwargs: Any,
    ) -> str:
        self._raise_if_schema_qualified(table)
        return cast(
            "str",
            super().visit_table(
                table,
                asfrom=asfrom,
                iscrud=iscrud,
                ashint=ashint,
                fromhints=fromhints,
                use_schema=use_schema,
                from_linter=from_linter,
                ambiguous_table_name_map=ambiguous_table_name_map,
                enclosing_alias=enclosing_alias,
                **kwargs,
            ),  # type: ignore[no-untyped-call]
        )

    @staticmethod
    def _is_true_onclause(onclause: Any) -> bool:
        node = onclause
        while node is not None:
            visit_name = getattr(node, "__visit_name__", None)
            if visit_name == "true":
                return True
            if visit_name in {"grouping", "unary"}:
                node = getattr(node, "element", None)
                continue
            return False
        return False

    @staticmethod
    def _validate_join_tree(join: Join) -> None:
        """Validate a Join tree against excel-dbapi constraints.

        Rejects:
        - Non-equality ON clauses (only col = col and AND-combined equalities allowed)
        - Non-column operands in ON (literals, functions, arithmetic)
        - Same-source ON comparisons (both operands from same table)
        """
        if isinstance(join.left, Join):
            ExcelCompiler._validate_join_tree(join.left)
        if isinstance(join.right, Join):
            ExcelCompiler._validate_join_tree(join.right)

        onclause = join.onclause
        if onclause is None:
            raise exc.CompileError("Excel dialect requires an ON clause for JOIN")

        if (
            ExcelCompiler._is_true_onclause(onclause)
            and not join.full
            and not join.isouter
        ):
            return

        def _collect_tables(from_clause: Any) -> set[Any]:
            if isinstance(from_clause, Join):
                return _collect_tables(from_clause.left) | _collect_tables(
                    from_clause.right
                )
            return {from_clause}

        left_tables = _collect_tables(join.left)
        right_tables = _collect_tables(join.right)

        def _check_on_clause(clause: Any) -> None:
            """Recursively validate that ON clause contains only cross-source column equalities."""
            visit_name = getattr(clause, "__visit_name__", None)
            if visit_name == "binary" and hasattr(clause, "operator"):
                if clause.operator is not operators.eq:
                    raise exc.CompileError(
                        "Excel dialect only supports '=' comparisons in JOIN ON clause"
                    )
                # Validate both operands are plain column references
                for operand in (clause.left, clause.right):
                    op_visit = getattr(operand, "__visit_name__", None)
                    if op_visit != "column":
                        raise exc.CompileError(
                            "Excel dialect only supports column references in JOIN ON clause"
                        )
                # Validate cross-source: one column from each side of the join
                left_tbl = getattr(clause.left, "table", None)
                right_tbl = getattr(clause.right, "table", None)
                left_from_left = left_tbl in left_tables
                left_from_right = left_tbl in right_tables
                right_from_left = right_tbl in left_tables
                right_from_right = right_tbl in right_tables
                cross = (left_from_left and right_from_right) or (
                    left_from_right and right_from_left
                )
                if not cross:
                    raise exc.CompileError(
                        "Excel dialect requires ON clause to compare columns from different join sources"
                    )
                return
            # AND-combined clause list (reject OR and other operators)
            if visit_name == "expression_clauselist":
                if clause.operator is not operators.and_:
                    raise exc.CompileError(
                        "Excel dialect does not support OR in JOIN ON clause"
                    )
                for sub_clause in clause.clauses:
                    _check_on_clause(sub_clause)
                return
            # Handle Grouping wrapper nodes
            if visit_name == "grouping":
                _check_on_clause(clause.element)
                return
            raise exc.CompileError(
                "Excel dialect only supports equality comparisons in JOIN ON clause"
            )

        _check_on_clause(onclause)

    def _setup_select_stack(
        self,
        select: Any,
        compile_state: Any,
        entry: Any,
        asfrom: Any,
        lateral: Any,
        compound_index: Any,
    ) -> Any:
        self._has_join = False
        setup_select_stack = cast("Callable[..., Any]", super()._setup_select_stack)
        froms = setup_select_stack(
            select,
            compile_state,
            entry,
            asfrom,
            lateral,
            compound_index,
        )
        for from_clause in froms:
            if isinstance(from_clause, Join):
                self._has_join = True
                break
        return froms

    def _reject_correlated_subquery(self, inner: Any) -> None:
        inner_tables = {
            from_obj.name
            for from_obj in inner.columns_clause_froms
            if hasattr(from_obj, "name")
        }
        # Collect tables referenced by nested subqueries so we skip them.
        # We only want to check correlations at the immediate level.
        nested_subquery_tables: set[str] = set()
        for elem in visitors.iterate(inner):
            visit_name = getattr(elem, "__visit_name__", None)
            if visit_name in ("select", "subquery") and elem is not inner:
                # Gather all table names from this nested scope
                for nested_elem in visitors.iterate(elem):
                    nested_tbl = getattr(nested_elem, "table", None)
                    if nested_tbl is not None and hasattr(nested_tbl, "name"):
                        nested_subquery_tables.add(nested_tbl.name)
        for elem in visitors.iterate(inner):
            table = getattr(elem, "table", None)
            if table is None or not hasattr(table, "name"):
                continue
            if table.name in nested_subquery_tables:
                continue
            if table.name not in inner_tables:
                raise exc.CompileError(
                    "Excel dialect does not support correlated subqueries"
                )

    def visit_function(
        self,
        func: Any,
        add_to_result_map: Any = None,
        **kwargs: Any,
    ) -> str:
        function_name = str(getattr(func, "name", "")).lower()
        if function_name not in _SUPPORTED_FUNCTIONS:
            raise exc.CompileError(
                f"Excel dialect does not support function: {function_name}"
            )
        result = str(
            super().visit_function(
                func,
                add_to_result_map=add_to_result_map,
                **kwargs,
            )
        )
        inner = result[result.index("(") + 1 : result.rindex(")")].strip()
        if not inner:
            raise exc.CompileError(
                f"Excel dialect does not support expression arguments in {function_name}()"
            )

        distinct_match = re.fullmatch(
            r"(?i)DISTINCT\s+([A-Za-z_][A-Za-z0-9_]*)",
            inner,
        )
        distinct_qualified_match = re.fullmatch(
            r"(?i)DISTINCT\s+([A-Za-z_][A-Za-z0-9_]*\.[A-Za-z_][A-Za-z0-9_]*)",
            inner,
        )
        if distinct_qualified_match:
            raise exc.CompileError(
                "Excel dialect does not support COUNT(DISTINCT table.col); "
                "use bare column names only"
            )
        if distinct_match:
            if function_name != "count":
                raise exc.CompileError(
                    f"Excel dialect does not support DISTINCT in {function_name}()"
                )
        elif inner != "*" and not re.fullmatch(
            r"[A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?",
            inner,
        ):
            raise exc.CompileError(
                f"Excel dialect does not support expression arguments in {function_name}()"
            )

        return result

    def visit_column(
        self,
        column: Any,
        add_to_result_map: Any = None,
        include_table: bool = True,
        result_map_targets: tuple[Any, ...] = (),
        ambiguous_table_name_map: MutableMapping[str, str] | None = None,
        **kwargs: Any,
    ) -> str:
        if kwargs.get("is_upsert_set"):
            table_name = getattr(getattr(column, "table", None), "name", None)
            use_table = table_name == "excluded"
        else:
            use_table = self._has_join
        return super().visit_column(
            column,
            add_to_result_map=add_to_result_map,
            include_table=use_table,
            result_map_targets=result_map_targets,
            ambiguous_table_name_map=ambiguous_table_name_map,
            **kwargs,
        )

    def visit_label(
        self,
        label: Any,
        add_to_result_map: Any = None,
        within_label_clause: bool = False,
        within_columns_clause: bool = False,
        render_label_as_label: Any = None,
        result_map_targets: Any = (),
        **kw: Any,
    ) -> str:
        """Emit AS <label> in SELECT columns for excel-dbapi alias support."""
        render_label_with_as = within_columns_clause and not within_label_clause

        if render_label_with_as:
            # Compute the label name for the result map
            if isinstance(label.name, elements._truncated_label):
                labelname = self._truncated_identifier("colident", label.name)
            else:
                labelname = label.name

            if add_to_result_map is not None:
                add_to_result_map(
                    labelname,
                    label.name,
                    (label, labelname) + label._alt_names + result_map_targets,
                    label.type,
                )

            # Emit the column WITH "AS <label>"
            return (
                str(
                    label.element._compiler_dispatch(
                        self,
                        within_columns_clause=True,
                        within_label_clause=True,
                        **kw,
                    )
                )
                + " AS "
                + self.preparer.format_label(label, labelname)
            )

        if render_label_as_label is label:
            if isinstance(label.name, elements._truncated_label):
                labelname = self._truncated_identifier("colident", label.name)
            else:
                labelname = label.name
            return self.preparer.format_label(label, labelname)

        return str(
            label.element._compiler_dispatch(self, within_columns_clause=False, **kw)
        )

    def visit_insert(
        self,
        insert_stmt: Any,
        *args: Any,
        **kw: Any,
    ) -> str:
        self._raise_if_schema_qualified(insert_stmt.table)
        visit_insert = cast("Callable[..., str]", super().visit_insert)
        return str(visit_insert(insert_stmt, *args, **kw))

    # ── Unsupported feature guards ─────────────────────────

    def visit_join(
        self,
        join: Any,
        asfrom: Any = False,
        from_linter: Any = None,
        **kwargs: Any,
    ) -> str:
        self._has_join = True
        self._validate_join_tree(join)

        if (
            join.onclause is not None
            and self._is_true_onclause(join.onclause)
            and not join.full
            and not join.isouter
        ):
            # Intentional cartesian product — register edges so the
            # FROM-linter treats both sides as connected (no spurious
            # SAWarning) while still detecting genuinely disconnected
            # FROM elements elsewhere in the statement.
            if from_linter:
                from_linter.edges.update(
                    itertools.product(
                        join.left._from_objects,
                        join.right._from_objects,
                    )
                )
            left = str(
                join.left._compiler_dispatch(
                    self,
                    asfrom=True,
                    from_linter=from_linter,
                    **kwargs,
                )
            )
            right = str(
                join.right._compiler_dispatch(
                    self,
                    asfrom=True,
                    from_linter=from_linter,
                    **kwargs,
                )
            )
            return left + " CROSS JOIN " + right

        visit_join = cast("Callable[..., str]", super().visit_join)
        return str(visit_join(join, asfrom=asfrom, from_linter=from_linter, **kwargs))

    def group_by_clause(self, select: Any, **kw: Any) -> str:
        return str(super().group_by_clause(select, **kw))  # type: ignore[no-untyped-call]

    def _compose_select_body(
        self,
        text: str,
        select: Any,
        compile_state: Any,
        inner_columns: Any,
        froms: Any,
        byfrom: Any,
        toplevel: bool,
        kwargs: Any,
    ) -> str:
        if self._has_join and select._distinct:
            raise exc.CompileError("Excel dialect does not support DISTINCT with JOIN")
        return str(
            super()._compose_select_body(  # type: ignore[no-untyped-call]
                text,
                select,
                compile_state,
                inner_columns,
                froms,
                byfrom,
                toplevel,
                kwargs,
            )
        )

    def limit_clause(self, select: Any, **kw: Any) -> str:
        text = ""
        if select._limit_clause is not None:
            text += " LIMIT " + self.process(select._limit_clause, **kw)
        if select._offset_clause is not None:
            text += " OFFSET " + self.process(select._offset_clause, **kw)
        return text

    def visit_cte(
        self,
        cte: Any,
        asfrom: bool = False,
        ashint: bool = False,
        fromhints: dict[Any, str] | None = None,
        visiting_cte: Any = None,
        from_linter: Any = None,
        cte_opts: Any = None,
        **kwargs: Any,
    ) -> str | None:
        raise exc.CompileError("Excel dialect does not support CTEs")

    def visit_subquery(self, subquery: Any, **kw: Any) -> str:
        if not self._in_in_clause:
            raise exc.CompileError(
                "Excel dialect only supports subqueries in WHERE ... IN"
            )

        if self._has_join:
            raise exc.CompileError(
                "Excel dialect does not support subqueries with JOIN"
            )

        if self._subquery_depth > 0:
            raise exc.CompileError("Excel dialect does not support nested subqueries")

        inner = getattr(subquery, "element", None)
        if inner is not None:
            # Reject subqueries that themselves contain a JOIN
            for from_clause in inner.get_final_froms():
                if isinstance(from_clause, Join):
                    raise exc.CompileError(
                        "Excel dialect does not support JOIN inside subqueries"
                    )
            self._reject_correlated_subquery(inner)

        self._subquery_depth += 1
        try:
            kw["literal_binds"] = True
            visit_subquery = cast("Callable[..., str]", super().visit_subquery)
            return str(visit_subquery(subquery, **kw))
        finally:
            self._subquery_depth -= 1

    def visit_grouping(self, grouping: Any, asfrom: bool = False, **kwargs: Any) -> str:
        element = getattr(grouping, "element", None)
        is_subquery_select = getattr(element, "__visit_name__", None) == "select"
        if is_subquery_select:
            if not self._in_in_clause:
                raise exc.CompileError(
                    "Excel dialect only supports subqueries in WHERE ... IN"
                )

            if self._has_join:
                raise exc.CompileError(
                    "Excel dialect does not support subqueries with JOIN"
                )

            if self._subquery_depth > 0:
                raise exc.CompileError(
                    "Excel dialect does not support nested subqueries"
                )

            assert element is not None  # guaranteed by is_subquery_select check
            # Reject subqueries that themselves contain a JOIN
            for from_clause in element.get_final_froms():
                if isinstance(from_clause, Join):
                    raise exc.CompileError(
                        "Excel dialect does not support JOIN inside subqueries"
                    )
            self._reject_correlated_subquery(element)
            kwargs["literal_binds"] = True

        if is_subquery_select:
            self._subquery_depth += 1
        try:
            visit_grouping = cast("Callable[..., str]", super().visit_grouping)
            return str(visit_grouping(grouping, asfrom=asfrom, **kwargs))
        finally:
            if is_subquery_select:
                self._subquery_depth -= 1

    def visit_binary(
        self,
        binary: Any,
        override_operator: Any = None,
        eager_grouping: bool = False,
        from_linter: Any = None,
        lateral_from_linter: Any = None,
        **kw: Any,
    ) -> str:
        binary_operator = override_operator or binary.operator
        in_context = (
            binary_operator is operators.in_op or binary_operator is operators.not_in_op
        )
        visit_binary = cast("Callable[..., str]", super().visit_binary)
        if not in_context:
            return str(
                visit_binary(
                    binary,
                    override_operator=override_operator,
                    eager_grouping=eager_grouping,
                    from_linter=from_linter,
                    lateral_from_linter=lateral_from_linter,
                    **kw,
                )
            )

        self._in_in_clause = True
        try:
            return str(
                visit_binary(
                    binary,
                    override_operator=override_operator,
                    eager_grouping=eager_grouping,
                    from_linter=from_linter,
                    lateral_from_linter=lateral_from_linter,
                    **kw,
                )
            )
        finally:
            self._in_in_clause = False

    def returning_clause(
        self,
        stmt: Any,
        returning_cols: Any,
        **kw: Any,
    ) -> str:
        raise exc.CompileError("Excel dialect does not support RETURNING")

    def for_update_clause(
        self,
        select: Any,
        **kw: Any,
    ) -> Literal[" FOR UPDATE"]:
        raise exc.CompileError("Excel dialect does not support SELECT ... FOR UPDATE")

    def visit_compound_select(
        self,
        cs: Any,
        asfrom: Any = False,
        compound_index: Any = None,
        **kwargs: Any,
    ) -> str:
        visit_compound_select = cast(
            "Callable[..., str]", super().visit_compound_select
        )
        sql = visit_compound_select(
            cs,
            asfrom=asfrom,
            compound_index=compound_index,
            **kwargs,
        )

        # SQLAlchemy parenthesizes branches that have their own ORDER BY,
        # e.g. "(SELECT id FROM t1 ORDER BY id DESC) UNION SELECT id FROM t2".
        # The excel-dbapi parser cannot handle leading '('.
        # Strip branch-local ORDER BY and its wrapping parentheses, but use
        # balanced-paren matching so inner parens (WHERE id IN (...),
        # function calls, IN (SELECT ...) subqueries) are NOT corrupted.
        sql = self._strip_compound_branch_parens(sql)
        return sql

    @staticmethod
    def _update_depth_quote_aware(sql: str, start: int, end: int, depth: int) -> int:
        """Update paren depth between *start* and *end*, skipping quoted strings."""
        in_quote = False
        quote_char = ""
        i = start
        while i < end:
            ch = sql[i]
            if in_quote:
                if ch == quote_char:
                    if i + 1 < end and sql[i + 1] == quote_char:
                        i += 2
                        continue
                    in_quote = False
            else:
                if ch in ("'", '"'):
                    in_quote = True
                    quote_char = ch
                elif ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
            i += 1
        return depth

    @staticmethod
    def _is_pos_in_quotes(sql: str, pos: int) -> bool:
        """Return True if *pos* is inside a quoted string literal."""
        in_quote = False
        quote_char = ""
        i = 0
        while i < pos:
            ch = sql[i]
            if in_quote:
                if ch == quote_char:
                    if i + 1 < len(sql) and sql[i + 1] == quote_char:
                        i += 2
                        continue
                    in_quote = False
            else:
                if ch in ("'", '"'):
                    in_quote = True
                    quote_char = ch
            i += 1
        return in_quote

    @staticmethod
    def _has_top_level_compound_op(sql: str) -> bool:
        """Return True if *sql* has a top-level UNION/INTERSECT/EXCEPT."""
        upper = sql.upper()
        depth = 0
        in_quote = False
        quote_char = ""
        i = 0
        length = len(sql)
        while i < length:
            ch = sql[i]
            if in_quote:
                if ch == quote_char:
                    if i + 1 < length and sql[i + 1] == quote_char:
                        i += 2
                        continue
                    in_quote = False
                i += 1
                continue
            if ch in ("'", '"'):
                in_quote = True
                quote_char = ch
                i += 1
                continue
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif depth == 0:
                for kw in ("UNION", "INTERSECT", "EXCEPT"):
                    kw_len = len(kw)
                    if upper[i : i + kw_len] == kw:
                        before_ok = i == 0 or not (
                            upper[i - 1].isalnum() or upper[i - 1] == "_"
                        )
                        after_pos = i + kw_len
                        after_ok = after_pos >= length or not (
                            upper[after_pos].isalnum() or upper[after_pos] == "_"
                        )
                        if before_ok and after_ok:
                            return True
            i += 1
        return False

    @staticmethod
    def _strip_compound_branch_parens(sql: str) -> str:
        """Strip outermost balanced parens that wrap compound branches.

        Only removes parentheses that wrap an entire compound branch,
        i.e. a ``(SELECT ...)`` group that appears at the very start of
        the SQL or immediately after a compound keyword (UNION, INTERSECT,
        EXCEPT, ALL).  ``IN (SELECT ...)`` subqueries inside a branch are
        preserved because they follow ``IN``, not a compound keyword.

        Branch-local ``ORDER BY`` is stripped in a parenthesis-aware way
        so that ``ORDER BY`` inside nested subqueries (e.g.
        ``IN (SELECT ... ORDER BY ...)``) is preserved.

        Raises ``exc.CompileError`` if a branch contains a nested compound
        operator (e.g. ``(SELECT ... UNION SELECT ...)``), which the
        excel-dbapi parser cannot handle.
        """
        import re

        branch_pos_re = re.compile(
            r"(?:^|(?:UNION|INTERSECT|EXCEPT|ALL))\s*$",
            re.IGNORECASE,
        )

        result: list[str] = []
        i = 0
        length = len(sql)

        while i < length:
            # Skip characters inside string literals at the top level.
            if sql[i] in ("'", '"'):
                quote_char = sql[i]
                j = i + 1
                while j < length:
                    if sql[j] == quote_char:
                        if j + 1 < length and sql[j + 1] == quote_char:
                            j += 2
                            continue
                        j += 1
                        break
                    j += 1
                result.append(sql[i:j])
                i = j
                continue

            if sql[i] == "(":
                # Find matching ')' using balanced counting, quote-aware.
                depth = 1
                in_quote = False
                qc = ""
                j = i + 1
                while j < length and depth > 0:
                    ch = sql[j]
                    if in_quote:
                        if ch == qc:
                            if j + 1 < length and sql[j + 1] == qc:
                                j += 2
                                continue
                            in_quote = False
                    else:
                        if ch in ("'", '"'):
                            in_quote = True
                            qc = ch
                        elif ch == "(":
                            depth += 1
                        elif ch == ")":
                            depth -= 1
                    j += 1
                inner = sql[i + 1 : j - 1].strip()

                prefix = "".join(result).rstrip()
                is_branch = (
                    inner.upper().startswith("SELECT")
                    and branch_pos_re.search(prefix) is not None
                )

                if is_branch:
                    # Reject grouped/nested compounds.
                    if ExcelCompiler._has_top_level_compound_op(inner):
                        raise exc.CompileError(
                            "Excel dialect does not support "
                            "grouped/nested compound queries. "
                            "Use flat chaining instead: "
                            "union(a, b).intersect(c) rather "
                            "than union(a, intersect(b, c))."
                        )
                    if ExcelCompiler._has_top_level_limit_offset(inner):
                        result.append("(" + inner + ")")
                    else:
                        inner = ExcelCompiler._strip_top_level_order_by(inner)
                        result.append(inner)
                else:
                    result.append(sql[i:j])
                i = j
            else:
                result.append(sql[i])
                i += 1

        return "".join(result)

    @staticmethod
    def _strip_top_level_order_by(sql: str) -> str:
        """Remove the last top-level ORDER BY from *sql*.

        Walks the string tracking parenthesis depth (quote-aware).  Only
        an ``ORDER BY`` token found at depth 0, outside string literals,
        is treated as branch-local and stripped.  Any subsequent ``LIMIT``
        or ``OFFSET`` at depth 0 is preserved so that
        ``ORDER BY x LIMIT n`` becomes ``LIMIT n``.
        """
        upper = sql.upper()
        depth = 0
        last_top_order: int | None = None
        search_start = 0
        while True:
            pos = upper.find("ORDER", search_start)
            if pos == -1:
                break
            # Skip if inside a quoted string.
            if ExcelCompiler._is_pos_in_quotes(sql, pos):
                search_start = pos + 5
                continue
            depth = ExcelCompiler._update_depth_quote_aware(
                sql, search_start, pos, depth
            )
            if depth == 0:
                after_order = pos + 5
                rest = upper[after_order:].lstrip()
                if (
                    rest.startswith("BY")
                    and (len(rest) <= 2 or not (rest[2].isalnum() or rest[2] == "_"))
                    and (
                        pos == 0
                        or not (upper[pos - 1].isalnum() or upper[pos - 1] == "_")
                    )
                ):
                    last_top_order = pos
            search_start = pos + 5
        if last_top_order is None:
            return sql

        order_end = len(sql)
        for keyword in ("LIMIT", "OFFSET"):
            od = 0
            scan2 = last_top_order
            while True:
                kp = upper.find(keyword, scan2)
                if kp == -1:
                    break
                if ExcelCompiler._is_pos_in_quotes(sql, kp):
                    scan2 = kp + len(keyword)
                    continue
                od = ExcelCompiler._update_depth_quote_aware(sql, scan2, kp, od)
                if od == 0 and (
                    kp == 0 or not (upper[kp - 1].isalnum() or upper[kp - 1] == "_")
                ):
                    after_kw = kp + len(keyword)
                    if after_kw >= len(upper) or not (
                        upper[after_kw].isalnum() or upper[after_kw] == "_"
                    ):
                        if kp < order_end:
                            order_end = kp
                        break
                scan2 = kp + len(keyword)

        before = sql[:last_top_order].rstrip()
        after = sql[order_end:]
        if after.strip():
            return before + " " + after.lstrip()
        return before

    @staticmethod
    def _has_top_level_limit_offset(sql: str) -> bool:
        """Return True if *sql* contains a top-level LIMIT or OFFSET."""
        upper = sql.upper()
        for keyword in ("LIMIT", "OFFSET"):
            depth = 0
            scan = 0
            while True:
                kp = upper.find(keyword, scan)
                if kp == -1:
                    break
                if ExcelCompiler._is_pos_in_quotes(sql, kp):
                    scan = kp + len(keyword)
                    continue
                depth = ExcelCompiler._update_depth_quote_aware(sql, scan, kp, depth)
                if depth == 0 and (
                    kp == 0 or not (upper[kp - 1].isalnum() or upper[kp - 1] == "_")
                ):
                    after_kw = kp + len(keyword)
                    if after_kw >= len(upper) or not (
                        upper[after_kw].isalnum() or upper[after_kw] == "_"
                    ):
                        return True
                scan = kp + len(keyword)
        return False

    def _on_conflict_target(self, clause: Any, **kw: Any) -> str:
        if clause.inferred_target_elements is not None:
            target_text = "({})".format(
                ", ".join(
                    (
                        self.preparer.quote(c)
                        if isinstance(c, str)
                        else self.process(c, include_table=False, use_schema=False)
                    )
                    for c in clause.inferred_target_elements
                )
            )
        else:
            target_text = ""

        return target_text

    @staticmethod
    def _rewrite_upsert_same_table_column(value: Any, insert_statement: Any) -> Any:
        if not isinstance(value, elements.ColumnElement):
            return value

        value_table = getattr(value, "table", None)
        insert_table = getattr(insert_statement, "table", None)
        if (
            value_table is None
            or insert_table is None
            or value_table is not insert_table
        ):
            return value

        excluded_columns = getattr(insert_statement, "excluded", None)
        value_key = getattr(value, "key", None)
        if excluded_columns is None or value_key is None:
            return value

        if value_key in excluded_columns:
            return excluded_columns[value_key]

        return value

    def visit_on_conflict_do_nothing(self, on_conflict: Any, **kw: Any) -> str:
        target_text = self._on_conflict_target(on_conflict, **kw)

        if target_text:
            return f"ON CONFLICT {target_text} DO NOTHING"
        else:
            return "ON CONFLICT DO NOTHING"

    def visit_on_conflict_do_update(self, on_conflict: Any, **kw: Any) -> str:
        clause = on_conflict

        target_text = self._on_conflict_target(on_conflict, **kw)

        action_set_ops: list[str] = []

        set_parameters = dict(clause.update_values_to_set)

        insert_statement = cast("Any", self.stack[-1]["selectable"])
        cols = insert_statement.table.c
        set_kw = dict(kw)
        set_kw.update(include_table=False, use_schema=False)
        for c in cols:
            col_key = c.key

            if col_key in set_parameters:
                value = set_parameters.pop(col_key)
            elif c in set_parameters:
                value = set_parameters.pop(c)
            else:
                continue

            value = self._rewrite_upsert_same_table_column(value, insert_statement)

            if coercions._is_literal(value):
                value = elements.BindParameter(None, value, type_=c.type)

            else:
                if isinstance(value, elements.BindParameter) and value.type._isnull:
                    value = value._clone()
                    value.type = c.type
            value_text = self.process(
                value.self_group(),
                is_upsert_set=True,
                **set_kw,
            )

            key_text = self.preparer.quote(c.name)
            action_set_ops.append(f"{key_text} = {value_text}")

        if set_parameters:
            from sqlalchemy import util as sa_util

            insert_table = getattr(self.current_executable, "table", None)
            insert_table_name = getattr(insert_table, "name", "<unknown>")

            # Reject foreign-table Column keys that would silently
            # retarget the assignment to the INSERT table
            for k in list(set_parameters):
                if (
                    hasattr(k, "table")
                    and hasattr(k.table, "name")
                    and k.table.name != insert_table_name
                ):
                    raise exc.CompileError(
                        f"Column '{k.key}' from table '{k.table.name}' cannot be used as "
                        f"an ON CONFLICT DO UPDATE SET target for table '{insert_table_name}'"
                    )

            additional_columns = ", ".join(f"'{column}'" for column in set_parameters)
            sa_util.warn(
                "Additional column names not matching any column keys in "
                f"table '{insert_table_name}': {additional_columns}"
            )
            for k, v in set_parameters.items():
                key_text = (
                    self.preparer.quote(k)
                    if isinstance(k, str)
                    else self.process(k, **set_kw)
                )
                value_text = self.process(
                    coercions.expect(roles.ExpressionElementRole, v),
                    is_upsert_set=True,
                    **set_kw,
                )
                action_set_ops.append(f"{key_text} = {value_text}")

        action_text = ", ".join(action_set_ops)

        return f"ON CONFLICT {target_text} DO UPDATE SET {action_text}"

    def visit_over(
        self,
        over: Any,
        **kwargs: Any,
    ) -> str:
        raise exc.CompileError("Excel dialect does not support window functions (OVER)")

    def visit_funcfilter(
        self,
        funcfilter: Any,
        **kwargs: Any,
    ) -> Any:
        raise exc.CompileError("Excel dialect does not support aggregate FILTER clause")
