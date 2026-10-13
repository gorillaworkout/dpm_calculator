"""Per-run upload limits shared by the web app, the chunked uploader and the CLI.

MAX_FILES counts every file in one Deal Segregator run: all Deals files
(MT5 Deals History and MT4 Raw Report share one field), the optional Client
Equity FX workbook, and the optional KVB MT4 account-type file.
A month of KVB daily files is 31 MT5 + 31 MT4 + FX + account types = 64.
"""
MAX_FILES = 70
