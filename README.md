# Find My Mutual Fund

A free, India-focused mutual fund sector exposure screener. Search and compare selected mutual fund schemes by disclosed portfolio allocation, using your own sector limits rather than a generic “best fund” ranking.

## Open the website

**Live screener:** https://bhardashubham7781.github.io/find-my-mutual-fund/

**How-to guide:** https://bhardashubham7781.github.io/find-my-mutual-fund/guides/compare-mutual-funds-by-sector-exposure.html

**Data and scoring methodology:** https://bhardashubham7781.github.io/find-my-mutual-fund/methodology.html

## What it does

- Screens mutual funds by disclosed sector exposure.
- Lets users exclude a sector or set maximum and minimum exposure limits.
- Supports 16 standardised sectors, including IT, Banking, Financial Services, Pharma & Healthcare, Auto, FMCG, Energy/Oil & Gas, Power, Metals & Mining, Capital Goods, Construction/Infra, Telecom, Consumer Durables, Realty, Chemicals and Others.
- Filters results by fund category and asset management company (AMC).
- Shows a rule-by-rule distance score so users can see how closely a scheme matches their selected conditions.

## Current data coverage

This is a **limited public beta** using selected scheme-matched monthly portfolios from September 2026. The current release contains 84 schemes: 38 from SBI Mutual Fund, 45 from ICICI Prudential Mutual Fund and 1 from HDFC Mutual Fund. It is not a complete list of Indian mutual funds, and other schemes or AMCs may be added only after their source files and scheme matches are reviewed.

Sector labels from different AMCs may be mapped to broader standard categories. Some source industries are assigned to **Others**, so the displayed category breakdown is an analytical aid rather than an official AMC classification. Portfolio disclosure data may be delayed.

## Use and limitations

A lower distance score means a closer match to the user's selected sector rules; it is not a rating of investment quality, risk or expected return. Verify scheme details and allocations against the relevant AMC's official monthly portfolio disclosure or factsheet before relying on them.

**Data is from AMC monthly portfolio disclosures and may be delayed. This is not investment advice.** Mutual fund investors should read scheme documents and consider their own objectives, time horizon and risk tolerance.

## Official source portals

- [SBI Mutual Fund portfolio disclosures](https://www.sbimf.com/portfoliospage_new)
- [HDFC Mutual Fund monthly portfolio disclosures](https://www.hdfcfund.com/statutory-disclosure/portfolio/monthly-portfolio)
- [Nippon India Mutual Fund factsheets and portfolio disclosures](https://mf.nipponindiaim.com/investor-service/downloads/factsheet-portfolio-and-other-disclosures/)
- [AMFI — Association of Mutual Funds in India](https://www.amfiindia.com/)

## Project repository

This repository contains the screener website, data-processing scripts, sector mapping and scheme review files. The public website is static and hosted by GitHub Pages.
