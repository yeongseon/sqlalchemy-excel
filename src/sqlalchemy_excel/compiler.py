"""SQL compiler for the Excel dialect.

Compiles SQLAlchemy expression trees into the SQL subset that
excel-dbapi's parser understands:

Supported:
    SELECT columns FROM table [WHERE ...] [GROUP BY ...] [HAVING ...] [ORDER BY col [ASC|DESC]] [LIMIT n] [OFFSET n]
    SELECT DISTINCT columns FROM table [WHERE ...] ...
    SELECT cols FROM t1 [INNER|LEFT|RIGHT] JOIN t2 ON ... { JOIN t3 ON ... } [WHERE ...] [ORDER BY ...] [LIMIT n] [OFFSET n]
    INSERT INTO table (cols) VALUES (vals), (vals2), ...
    INSERT INTO table (cols) SELECT cols FROM source [WHERE ...]
    UPDATE table SET col=val [WHERE ...]
    DELETE FROM table [WHERE ...]

    Rejected (raises CompileError):
    CTEs, window functions, RETURNING, FOR UPDATE, NOT IN,
    FULL OUTER JOIN, CROSS JOIN, NATURAL JOIN, non-equality/OR/non-column ON clauses

    Partially supported:
    non-correlated subqueries in WHERE ... IN (SELECT single_col FROM table [WHERE ...])

For single-table queries, excel-dbapi expects unquoted, unprefixed column names:
    SELECT id, name FROM users          (correct)
    SELECT users.id, users.name FROM users  (WRONG — parser rejects)

For JOIN queries, table-qualified column names are required:
    SELECT users.id, orders.user_id FROM users JOIN orders ON users.id = orders.user_id

The compiler detects JOIN context and switches between the two modes automatically.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any, Literal, cast

from sqlalchemy import exc
from sqlalchemy.sql import compiler, dml, elements, operators, visitors
from sqlalchemy.sql.expression import Join

if TYPE_CHECKING:
    from collections.abc import Callable, MutableMapping


_SUPPORTED_FUNCTIONS = {"count", "sum", "avg", "min", "max"}


class ExcelIdentifierPreparer(compiler.IdentifierPreparer):
    """Identifier preparer that never quotes identifiers.

    excel-dbapi's parser expects bare column names without quotes
    or table prefixes.
    """

    def __init__(self, dialect: Any) -> None:
        super().__init__(dialect, initial_quote="", final_quote="")
        self.reserved_words = set()

    def quote_identifier(self, value: str) -> str:
        return value

    def quote(self, ident: str, force: Any = None) -> str:
        return ident


class ExcelCompiler(compiler.SQLCompiler):
    """Compiles SQLAlchemy SQL expressions for excel-dbapi."""

    _in_in_clause: bool = False
    _subquery_depth: int = 0
    _has_join: bool = False

    @staticmethod
    def _validate_join_tree(join: Join) -> None:
        """Validate a Join tree against excel-dbapi constraints.

        Rejects:
        - FULL OUTER JOIN
        - Non-equality ON clauses (only col = col and AND-combined equalities allowed)
        - Non-column operands in ON (literals, functions, arithmetic)
        - Same-source ON comparisons (both operands from same table)
        """
        if isinstance(join.left, Join):
            ExcelCompiler._validate_join_tree(join.left)
        if isinstance(join.right, Join):
            ExcelCompiler._validate_join_tree(join.right)

        # 2. Reject FULL OUTER JOIN
        if join.full:
            raise exc.CompileError(
                "Excel dialect does not support FULL OUTER JOIN"
            )

        # 3. Validate ON clause structure: only equality comparisons allowed
        onclause = join.onclause
        if onclause is None:
            raise exc.CompileError(
                "Excel dialect requires an ON clause for JOIN"
            )

        def _collect_tables(from_clause: Any) -> set[Any]:
            if isinstance(from_clause, Join):
                return _collect_tables(from_clause.left) | _collect_tables(from_clause.right)
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
            # Anything else (true(), unary, etc.) is unsupported
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
        setup_select_stack = cast(
            "Callable[..., Any]", super()._setup_select_stack
        )
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

        if inner != "*" and not re.fullmatch(
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
        """Override to conditionally include table prefix in column references.

        For single-table queries: SELECT id, name FROM users
        For JOIN queries: SELECT users.id, orders.amount FROM users JOIN orders ...
        """
        return super().visit_column(
            column,
            add_to_result_map=add_to_result_map,
            include_table=self._has_join,
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
        """Override to never emit AS <label> in SELECT columns.

        excel-dbapi's parser does not understand column aliases.
        We still register the result map so SQLAlchemy can map
        result columns back to ORM attributes by position.
        """
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

            # Emit the column WITHOUT "AS <label>"
            return str(
                label.element._compiler_dispatch(
                    self,
                    within_columns_clause=True,
                    within_label_clause=True,
                    **kw,
                )
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
        visit_join = cast("Callable[..., str]", super().visit_join)
        return str(
            visit_join(join, asfrom=asfrom, from_linter=from_linter, **kwargs)
        )

    def group_by_clause(self, select: Any, **kw: Any) -> str:
        if self._has_join and select._group_by_clauses:
            raise exc.CompileError(
                "Excel dialect does not support GROUP BY with JOIN"
            )
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
        if self._has_join:
            # Reject DISTINCT + JOIN
            if select._distinct:
                raise exc.CompileError(
                    "Excel dialect does not support DISTINCT with JOIN"
                )
            # Reject HAVING + JOIN
            if select._having_criteria:
                raise exc.CompileError(
                    "Excel dialect does not support HAVING with JOIN"
                )
            # Reject aggregates + JOIN
            for col_elem in inner_columns:
                col_text = str(col_elem)
                for agg in _SUPPORTED_FUNCTIONS:
                    if col_text.upper().startswith(agg.upper() + "("):
                        raise exc.CompileError(
                            "Excel dialect does not support aggregate functions with JOIN"
                        )
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
            raise exc.CompileError(
                "Excel dialect does not support nested subqueries"
            )

        if isinstance(self.statement, dml.UpdateBase):
            raise exc.CompileError(
                "Excel dialect does not support subqueries in UPDATE/DELETE"
            )

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

    def visit_grouping(
        self, grouping: Any, asfrom: bool = False, **kwargs: Any
    ) -> str:
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

            if isinstance(self.statement, dml.UpdateBase):
                raise exc.CompileError(
                    "Excel dialect does not support subqueries in UPDATE/DELETE"
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
        in_context = binary_operator is operators.in_op
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

    def visit_not_in_op_binary(self, binary: Any, operator: Any, **kw: Any) -> str:
        raise exc.CompileError("Excel dialect does not support NOT IN")

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
        """
        import re

        # Regex to detect compound keyword or start-of-string before '('.
        _BRANCH_POS_RE = re.compile(
            r"(?:^|(?:UNION|INTERSECT|EXCEPT|ALL))\s*$",
            re.IGNORECASE,
        )

        result: list[str] = []
        i = 0
        length = len(sql)

        while i < length:
            if sql[i] == "(":
                # Find matching ')' using balanced counting.
                depth = 1
                j = i + 1
                while j < length and depth > 0:
                    if sql[j] == "(":
                        depth += 1
                    elif sql[j] == ")":
                        depth -= 1
                    j += 1
                # j is one past the matching ')'.
                inner = sql[i + 1 : j - 1].strip()

                # Only strip if (a) content starts with SELECT AND
                # (b) the '(' is at a compound-branch position (start
                # of string or after UNION/INTERSECT/EXCEPT/ALL).
                prefix = "".join(result).rstrip()
                is_branch = (
                    inner.upper().startswith("SELECT")
                    and _BRANCH_POS_RE.search(prefix) is not None
                )

                if is_branch:
                    # If the branch has LIMIT/OFFSET, the parser
                    # can handle the full form inside parens
                    # (ORDER BY + LIMIT/OFFSET), so keep parens
                    # intact to preserve both ordering and limiting
                    # semantics.
                    if ExcelCompiler._has_top_level_limit_offset(inner):
                        result.append("(" + inner + ")")
                    else:
                        # No LIMIT/OFFSET — strip branch-local
                        # ORDER BY (which is semantically meaningless
                        # in compound queries without LIMIT) and
                        # remove the wrapper parens.
                        inner = ExcelCompiler._strip_top_level_order_by(
                            inner
                        )
                        result.append(inner)
                else:
                    # Not a branch wrapper — preserve as-is.
                    result.append(sql[i:j])
                i = j
            else:
                result.append(sql[i])
                i += 1

        return "".join(result)

    @staticmethod
    def _strip_top_level_order_by(sql: str) -> str:
        """Remove the last top-level ORDER BY from *sql*.

        Walks the string tracking parenthesis depth.  Only an ``ORDER BY``
        token found at depth 0 is treated as branch-local and stripped.
        Any subsequent ``LIMIT`` or ``OFFSET`` at depth 0 is preserved so
        that ``ORDER BY x LIMIT n`` becomes ``LIMIT n``.
        An ``ORDER BY`` inside any parenthesized group (e.g. a subquery)
        is left untouched.
        """
        upper = sql.upper()
        depth = 0
        last_top_order: int | None = None
        search_start = 0
        # Scan for top-level ORDER BY positions (paren depth == 0).
        while True:
            pos = upper.find("ORDER", search_start)
            if pos == -1:
                break
            # Update paren depth up to this position.
            for k in range(search_start, pos):
                if sql[k] == "(":
                    depth += 1
                elif sql[k] == ")":
                    depth -= 1
            if depth == 0:
                # Verify it's ORDER BY (not part of an identifier).
                after_order = pos + 5
                rest = upper[after_order:].lstrip()
                if rest.startswith("BY"):
                    if pos == 0 or not upper[pos - 1].isalnum():
                        last_top_order = pos
            search_start = pos + 5
        if last_top_order is None:
            return sql

        # Find where the ORDER BY clause ends: at the next top-level
        # LIMIT or OFFSET keyword, or at end-of-string.
        order_end = len(sql)
        od = 0  # paren depth from last_top_order onward
        scan = last_top_order
        for keyword in ("LIMIT", "OFFSET"):
            od = 0
            scan2 = last_top_order
            while True:
                kp = upper.find(keyword, scan2)
                if kp == -1:
                    break
                # Update paren depth up to this position.
                for k2 in range(scan2, kp):
                    if sql[k2] == "(":
                        od += 1
                    elif sql[k2] == ")":
                        od -= 1
                if od == 0:
                    # Verify it's not part of an identifier.
                    if kp == 0 or not upper[kp - 1].isalnum():
                        after_kw = kp + len(keyword)
                        if after_kw >= len(upper) or not upper[after_kw].isalnum():
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
                for k in range(scan, kp):
                    if sql[k] == "(":
                        depth += 1
                    elif sql[k] == ")":
                        depth -= 1
                if depth == 0:
                    # Not part of an identifier.
                    if kp == 0 or not upper[kp - 1].isalnum():
                        after_kw = kp + len(keyword)
                        if after_kw >= len(upper) or not upper[after_kw].isalnum():
                            return True
                scan = kp + len(keyword)
        return False

    def visit_over(
        self,
        over: Any,
        **kwargs: Any,
    ) -> str:
        raise exc.CompileError(
            "Excel dialect does not support window functions (OVER)"
        )

    def visit_funcfilter(
        self,
        funcfilter: Any,
        **kwargs: Any,
    ) -> Any:
        raise exc.CompileError(
            "Excel dialect does not support aggregate FILTER clause"
        )
