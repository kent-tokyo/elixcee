# 0.21.0 Streaming Reader Hardening

The streaming reader now tokenizes the worksheet entry at XML tag boundaries rather than
searching independently inside fixed-size byte chunks. A `<row` or `</row>` split across the
64 KiB decompression buffer is therefore handled correctly. Rows continue to travel through
a bounded channel, preserving the forward-only API and avoiding workbook-wide materialization.

This release keeps the intentionally narrow stream contract: values only, no VBA, formulas,
styles, merges, charts, or arbitrary cell mutation. The append-only writer API remains
compatible with 0.20.0; constant-memory writer output is a follow-up scope.
