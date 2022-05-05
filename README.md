# Pronto SPL Processor

Transform supplier price lists to Pronto SPL import format.

## System Requirements

- python >=3.10
- openpyxl

## Usage

```
python -m spl_proc <reader>
```

## Available readers

| Name  | Used For                                       |
| :---- | :--------------------------------------------- |
| `BRO` | Brother pricelist (converted from PDF to XLSX) |
| `CSS` | Creative School Supply Company pricelist       |
| `DYN` | Dynamic Supplies pricelist                     |
