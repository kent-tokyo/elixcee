"""A tiny, dependency-free elixcee playground for first-time users."""

from __future__ import annotations

import argparse
import json

import elixcee


VBA = """
Sub AddBonus()
    Cells(2, 2).Value = Cells(2, 1).Value + 5
End Sub
"""


def main() -> None:
    parser = argparse.ArgumentParser(description="Try elixcee in a few lines")
    parser.add_argument("--output", help="optional .xlsx output path")
    args = parser.parse_args()

    vm = elixcee.Vm()
    vm.set_cell(1, 1, "商品")
    vm.set_cell(1, 2, "金額")
    vm.set_cell(2, 1, 100)
    vm.set_cell(3, 1, 250)
    vm.set_cell_formula(4, 1, "=SUM(A2:A3)")
    vm.run(VBA, "AddBonus")

    if args.output:
        vm.save_workbook(args.output)

    print(
        json.dumps(
            {
                "A1": vm.get_cell(1, 1),
                "A4_total": vm.get_cell(4, 1),
                "B2_bonus": vm.get_cell(2, 2),
                "output": args.output,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
