# Move Soldout Products to Archived Pages

This app takes a **Matrixify** export file with active products, filter those that was last updated more than
XY days ago and are currently hidden and deletes them. It also creates pages with the same handle and the same
content as deleted products and sets redirects.

**Output is an Excel file to be imported by Matrixify.**

Necessary columns in source file are:
`ID`, `Handle`, `Command`, `Title`, `Body HTML`, `Tags`, `Variant SKU` and all `Variant Metafields`.

The name of the source file should contain of keyword `source` and the locale code, e.g. `source_cz.xlsx` or `source_sk.xlsx`.

<br>

### Proceed with caution!

<br>

To run this app you need a `secrets.py` file in the app folder with ShopPIM credentials:

```python
HOST = "<IP ADDRESS OF SQL DB>"
USER = "<USERNAME TO SQL DB>"
PASS = "<PASSWORD TO SQL DB>"
```