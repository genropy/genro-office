#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Annual report with @component — reusable monthly blocks.

Demonstrates how a single ``month_block`` component is reused 12 times
in a loop, each with a different data prefix. The same structure
(header + metrics + profit + spacer) is repeated for every month,
but pointer paths resolve to different data.

Architecture:
    ReportBuilder  — extends ExcelBuilder, adds @component definitions
    ReportCompiler — extends ExcelCompiler, flattens component wrappers in sheets
    AnnualReport   — extends ExcelApp, uses month_block in a loop
"""

from __future__ import annotations

import random
from typing import Any

from genro_bag import Bag
from genro_builders.builder import component, element

from genro_office import ExcelApp
from genro_office.builders.excel_builder import ExcelBuilder
from genro_office.compilers.excel_compiler import ExcelCompiler

# ---------------------------------------------------------------------------
# Builder: vocabulary with components
# ---------------------------------------------------------------------------

class ReportBuilder(ExcelBuilder):
    """ExcelBuilder extended with report-specific components."""

    # Redefine sheet to accept component tags as children
    @element(sub_tags="row,merge,chart,month_block", parent_tags="workbook")
    def sheet(
        self,
        name: str = "Sheet1",
        freeze_panes: str | None = None,
        autofilter: str | None = None,
    ) -> None:
        ...

    @component(sub_tags="", parent_tags="sheet")
    def month_block(
        self,
        comp: Any,
        month_name: str = "January",
        data_prefix: str = "month.jan",
        **kwargs: Any,  # noqa: ARG002
    ) -> None:
        """Reusable monthly data block: header + metrics + profit + spacer.

        Args:
            comp: The component's internal Bag.
            month_name: Display name for the month header.
            data_prefix: Data path prefix for ^pointer resolution.
        """
        # Header row
        header = comp.row(height=22.0)
        header.cell(
            content=month_name, bold=True,
            bg_color="4472C4", font_color="FFFFFF", width=15.0,
        )
        header.cell(
            content="Value", bold=True,
            bg_color="4472C4", font_color="FFFFFF", width=12.0,
        )
        header.cell(
            content="vs Budget", bold=True,
            bg_color="4472C4", font_color="FFFFFF", width=12.0,
        )

        # Metric rows
        for metric in ("revenue", "expenses", "headcount"):
            row = comp.row()
            row.cell(content=metric.title())
            row.cell(content=f"^{data_prefix}?{metric}", number_format="#,##0")
            row.cell(content=f"^{data_prefix}?{metric}_vs_budget")

        # Profit row (bold)
        profit = comp.row()
        profit.cell(content="Profit", bold=True, bg_color="E2EFDA")
        profit.cell(
            content=f"^{data_prefix}?profit", bold=True,
            number_format="#,##0", bg_color="E2EFDA",
        )
        profit.cell(content=f"^{data_prefix}?profit_vs_budget", bold=True, bg_color="E2EFDA")

        # Spacer row
        comp.row()


# ---------------------------------------------------------------------------
# Compiler: component transparency for sheets
# ---------------------------------------------------------------------------

class ReportCompiler(ExcelCompiler):
    """ExcelCompiler that flattens component wrappers inside sheets."""

    def _build_sheet(self, node: Any, wb: Any) -> None:
        """Build sheet node, flattening component wrappers."""
        name = node.attr.get("name", "Sheet1")
        ws = wb.create_sheet(title=str(name))

        freeze_panes = node.attr.get("freeze_panes")
        if freeze_panes:
            ws.freeze_panes = str(freeze_panes)

        autofilter = node.attr.get("autofilter")
        if autofilter:
            ws.auto_filter.ref = str(autofilter)

        if isinstance(node.value, Bag):
            merge_nodes: list[Any] = []
            chart_nodes: list[Any] = []

            row_idx = 1
            for child_node in self._iter_sheet_children(node.value):
                if child_node.node_tag == "row":
                    self._build_row(child_node, ws, row_idx)
                    row_idx += 1
                elif child_node.node_tag == "merge":
                    merge_nodes.append(child_node)
                elif child_node.node_tag == "chart":
                    chart_nodes.append(child_node)

            for merge_node in merge_nodes:
                self._build_merge(merge_node, ws)

            for chart_node in chart_nodes:
                self._build_chart(chart_node, ws)

    def _iter_sheet_children(self, bag: Bag) -> Any:
        """Yield row/merge/chart nodes, flattening component wrappers."""
        for node in bag:
            tag = node.node_tag or ""
            if tag in ("row", "merge", "chart"):
                yield node
            elif isinstance(node.value, Bag):
                yield from self._iter_sheet_children(node.value)


# Link compiler to builder
ReportBuilder._compiler_class = ReportCompiler


# ---------------------------------------------------------------------------
# App: the annual report template
# ---------------------------------------------------------------------------

MONTHS = [
    ("January", "jan"),
    ("February", "feb"),
    ("March", "mar"),
    ("April", "apr"),
    ("May", "may"),
    ("June", "jun"),
    ("July", "jul"),
    ("August", "aug"),
    ("September", "sep"),
    ("October", "oct"),
    ("November", "nov"),
    ("December", "dec"),
]


class AnnualReport(ExcelApp):
    """Annual financial report using reusable monthly block components."""

    def __init__(self) -> None:
        self.builder = self.set_builder("main", ReportBuilder)

    def main(self, source: Any) -> None:
        wb = source.workbook()
        sheet = wb.sheet(name="Annual Report 2025")

        # 12 identical blocks, each with different data prefix
        for month_name, key in MONTHS:
            sheet.month_block(month_name=month_name, data_prefix=f"month.{key}")


def _generate_sample_data() -> list[tuple[str, str, dict[str, Any]]]:
    """Generate realistic sample data for 12 months."""
    random.seed(42)
    data = []
    base_revenue = 100000
    base_expenses = 70000
    base_headcount = 50

    for month_name, key in MONTHS:
        revenue = base_revenue + random.randint(-10000, 20000)
        expenses = base_expenses + random.randint(-5000, 10000)
        headcount = base_headcount + random.randint(-3, 5)
        profit = revenue - expenses

        data.append((
            month_name,
            key,
            {
                "revenue": revenue,
                "expenses": expenses,
                "headcount": headcount,
                "profit": profit,
                "revenue_vs_budget": f"+{random.randint(1, 15)}%",
                "expenses_vs_budget": f"{random.choice(['+', '-'])}{random.randint(1, 10)}%",
                "headcount_vs_budget": f"{random.choice(['+', '-'])}{random.randint(0, 3)}",
                "profit_vs_budget": f"+{random.randint(1, 20)}%",
            },
        ))

        base_revenue += random.randint(-2000, 5000)
        base_expenses += random.randint(-1000, 3000)
        base_headcount += random.randint(0, 2)

    return data


if __name__ == "__main__":
    report = AnnualReport()

    for _month_name, key, month_data in _generate_sample_data():
        report.data.set_item(f"month.{key}", "", **month_data)

    report.build()
    report.save("output.xlsx")
    print("Created: output.xlsx")
