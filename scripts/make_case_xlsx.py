from __future__ import annotations

import argparse

from eia_gen.services.xlsx.case_template import write_case_template_xlsx


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--out", default="templates/case_template.xlsx")
    args = p.parse_args()
    out = write_case_template_xlsx(args.out)
    print(f"wrote: {out}")


if __name__ == "__main__":
    main()

