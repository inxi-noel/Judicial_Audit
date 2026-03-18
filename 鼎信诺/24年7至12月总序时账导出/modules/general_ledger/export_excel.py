from modules.general_ledger.ledger_runtime import (
    build_expected_export_filename,
    export_multiyear_general_ledger,
)


if __name__ == "__main__":
    print(export_multiyear_general_ledger("043\u5408\u80a5\u9e3f\u67cf\u5730\u4ea7\u5f00\u53d1\u6709\u9650\u8d23\u4efb\u516c\u53f8", 7))
