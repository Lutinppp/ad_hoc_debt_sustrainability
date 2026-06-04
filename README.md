# ad_hoc_debt_sustrainability

IMF API-based replacement for the legacy VBA debt-sustainability workflow.

This implementation generates one Excel workbook in one run for:
- EZ (Euro Area)
- FR (France)
- DE (Germany)
- IT (Italy)
- ES (Spain)
- US (United States)
- UK (United Kingdom)

## What It Produces

For each geography sheet, the script writes:
- A top zero-flow block with:
	- Gross debt (% GDP)
	- Projected debt stock (% GDP)
	- Primary net lending/borrowing (% GDP)
	- Nominal GDP growth (%)
	- Other identified flows set to 0
	- Implied effective rate (%)
	- Debt-stabilizing differential
	- Primary balance minus differential
- A second calibrated-flow block with the same rows, except row 8 is computed other identified flows derived as the residual needed to align the recursive debt identity with the IMF debt series

Column B is a hidden lag year (same behavior as the VBA macro logic).

## Install

```bash
pip install -r requirements.txt
```

## IMF API Credentials Module

Edit the module below to paste your IMF API credentials if your IMF endpoint requires auth:

- `src/pbs_imf/config.py`

Fields to edit:
- `api_key`
- `bearer_token`

Environment variables can also be used and will override file values:
- `IMF_API_KEY`
- `IMF_BEARER_TOKEN`
- `IMF_BASE_URL`
- `IMF_TIMEOUT_SECONDS`

## Run

```bash
python run_pbs_imf.py --start-year 2020 --end-year 2030 --output pbs_imf_output.xlsx
```

## Export Raw IMF Series

To export the underlying IMF series from a single standalone Python file with one sheet per geography, run:

```bash
python run_imf_series_export.py --start-year 2020 --end-year 2030 --output imf_series_export.xlsx
```

Each country sheet contains:
- Year
- Gross debt (% GDP)
- Primary balance (% GDP)
- GDP level (`NGDP` or fallback `NGDPD`, sourced from IMF SDMX WEO; falls back to DataMapper when SDMX is empty for a country)
- Nominal GDP growth (%)

## Data Notes

- Debt indicator: `GGXWDG_NGDP`
- Primary balance indicator: `GGXONLB_G01_GDP_PT`
- GDP level in the standalone exporter: pulled from IMF SDMX WEO, tries `NGDP` first, then falls back to `NGDPD`; if SDMX has no GDP values for a country, it falls back to DataMapper.

The DataMapper endpoint currently used is:
- `https://www.imf.org/external/datamapper/api/v1`