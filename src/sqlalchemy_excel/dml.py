"""DML constructs for the Excel dialect -- INSERT ON CONFLICT support."""

from __future__ import annotations

from typing import Any, Optional, Union, cast

from sqlalchemy import util
from sqlalchemy.sql import coercions, roles, schema
from sqlalchemy.sql._typing import _DMLTableArgument
from sqlalchemy.sql.base import (
    ColumnCollection,
    ReadOnlyColumnCollection,
    _exclusive_against,
    _generative,
)
from sqlalchemy.sql.dml import Insert as StandardInsert
from sqlalchemy.sql.elements import (
    ClauseElement,
    ColumnElement,
    KeyedColumnElement,
    TextClause,
)
from sqlalchemy.sql.expression import alias
from sqlalchemy.util.typing import Self

__all__ = ("Insert", "insert")

_OnConflictIndexElementsT = Optional[list[Union[str, schema.Column[Any]]]]
_OnConflictIndexWhereT = Optional[Union[ColumnElement[Any], TextClause]]
_OnConflictSetDictT = dict[Union[schema.Column[Any], str], Any]
_OnConflictSetT = Optional[Union[_OnConflictSetDictT, ColumnCollection[Any, Any]]]
_OnConflictWhereT = Optional[Union[ColumnElement[Any], TextClause]]


def insert(table: _DMLTableArgument) -> Insert:
    """Construct an Excel-specific variant :class:`_excel.Insert` construct."""
    return Insert(table)


class Insert(StandardInsert):
    """Excel-specific implementation of INSERT.

    Adds methods for ON CONFLICT (UPSERT) support.
    """

    stringify_dialect = "excel"
    inherit_cache = False

    @util.memoized_property
    def excluded(self) -> ReadOnlyColumnCollection[str, KeyedColumnElement[Any]]:
        """Provide the ``excluded`` namespace for an ON CONFLICT statement."""
        return alias(self.table, name="excluded").columns

    _on_conflict_exclusive = _exclusive_against(
        "_post_values_clause",
        msgs={
            "_post_values_clause": "This Insert construct already has "
            "an ON CONFLICT clause established"
        },
    )

    @_generative
    @_on_conflict_exclusive
    def on_conflict_do_update(
        self,
        index_elements: _OnConflictIndexElementsT = None,
        index_where: _OnConflictIndexWhereT = None,
        set_: _OnConflictSetT = None,
        where: _OnConflictWhereT = None,
    ) -> Self:
        """Specifies a DO UPDATE SET action for ON CONFLICT clause."""
        if not index_elements:
            raise ValueError(
                "Excel dialect requires index_elements for ON CONFLICT clause"
            )
        if index_where is not None:
            raise ValueError("Excel dialect does not support index_where in ON CONFLICT")
        if where is not None:
            raise ValueError(
                "Excel dialect does not support WHERE clause in ON CONFLICT DO UPDATE"
            )
        self._post_values_clause = OnConflictDoUpdate(
            index_elements, index_where, set_, where
        )
        return self

    @_generative
    @_on_conflict_exclusive
    def on_conflict_do_nothing(
        self,
        index_elements: _OnConflictIndexElementsT = None,
        index_where: _OnConflictIndexWhereT = None,
    ) -> Self:
        """Specifies a DO NOTHING action for ON CONFLICT clause."""
        if not index_elements:
            raise ValueError(
                "Excel dialect requires index_elements for ON CONFLICT clause"
            )
        if index_where is not None:
            raise ValueError("Excel dialect does not support index_where in ON CONFLICT")
        self._post_values_clause = OnConflictDoNothing(index_elements, index_where)
        return self


class OnConflictClause(ClauseElement):
    stringify_dialect = "excel"

    inferred_target_elements: Optional[list[Union[str, schema.Column[Any]]]]
    inferred_target_whereclause: Optional[Union[ColumnElement[Any], TextClause]]

    def __init__(
        self,
        index_elements: _OnConflictIndexElementsT = None,
        index_where: _OnConflictIndexWhereT = None,
    ):
        if index_elements is not None:
            self.inferred_target_elements = [
                coercions.expect(roles.DDLConstraintColumnRole, column)
                for column in index_elements
            ]
            self.inferred_target_whereclause = (
                coercions.expect(roles.WhereHavingRole, index_where)
                if index_where is not None
                else None
            )
        else:
            self.inferred_target_elements = self.inferred_target_whereclause = None


class OnConflictDoNothing(OnConflictClause):
    __visit_name__ = "on_conflict_do_nothing"


class OnConflictDoUpdate(OnConflictClause):
    __visit_name__ = "on_conflict_do_update"

    update_values_to_set: list[tuple[Union[schema.Column[Any], str], Any]]
    update_whereclause: Optional[ColumnElement[Any]]

    def __init__(
        self,
        index_elements: _OnConflictIndexElementsT = None,
        index_where: _OnConflictIndexWhereT = None,
        set_: _OnConflictSetT = None,
        where: _OnConflictWhereT = None,
    ):
        super().__init__(index_elements=index_elements, index_where=index_where)

        if isinstance(set_, dict):
            if not set_:
                raise ValueError("set parameter dictionary must not be empty")
            set_dict = set_
        elif isinstance(set_, ColumnCollection):
            set_dict = {column: column for column in set_}
            if not set_dict:
                raise ValueError("set parameter dictionary must not be empty")
        else:
            raise ValueError(
                "set parameter must be a non-empty dictionary "
                "or a ColumnCollection such as the `.c.` collection "
                "of a Table object"
            )
        self.update_values_to_set = [
            (coercions.expect(roles.DMLColumnRole, key), value)
            for key, value in set_dict.items()
        ]
        self.update_whereclause = (
            coercions.expect(roles.WhereHavingRole, where)
            if where is not None
            else None
        )
