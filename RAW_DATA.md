# Donnees financieres brutes

## Perimetre

- Source principale: SEC companyfacts + Form 10-K SEC
- Annee comparee: `2025`
- Valeurs de bilan conservees pour `2024` et `2025` afin de calculer les moyennes demandees par le cours
- Unites: USD

## Visa

| Metric | 2024 | 2025 | Source |
|---|---:|---:|---|
| Revenue | 35,926,000,000 | 40,000,000,000 | SEC companyfacts, `RevenueFromContractWithCustomerExcludingAssessedTax` |
| Net income | 19,743,000,000 | 20,058,000,000 | SEC companyfacts, `NetIncomeLoss` |
| Total assets | 94,511,000,000 | 99,627,000,000 | SEC companyfacts, `Assets` |
| Current assets | 34,033,000,000 | 37,766,000,000 | SEC companyfacts, `AssetsCurrent` |
| Cash and cash equivalents | 11,975,000,000 | 17,164,000,000 | SEC companyfacts, `CashAndCashEquivalentsAtCarryingValue` |
| Accounts receivable, net | 2,561,000,000 | 3,126,000,000 | SEC companyfacts, `AccountsReceivableNetCurrent` |
| Current liabilities | 26,517,000,000 | 35,048,000,000 | SEC companyfacts, `LiabilitiesCurrent` |
| Total liabilities | 55,374,000,000 | 61,718,000,000 | SEC companyfacts, `Liabilities` |
| Total equity | 39,137,000,000 | 37,909,000,000 | SEC companyfacts, `StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest` |
| Long-term debt current portion | 0 | 5,569,000,000 | SEC companyfacts, `LongTermDebtCurrent` |
| Long-term debt noncurrent | 20,836,000,000 | 19,602,000,000 | SEC companyfacts, `LongTermDebtNoncurrent` |
| Total debt used for project | 20,836,000,000 | 25,171,000,000 | Derived = current portion + noncurrent |
| PPE, net | 3,824,000,000 | 4,236,000,000 | SEC companyfacts, `PropertyPlantAndEquipmentNet` |
| Goodwill | 18,941,000,000 | 19,879,000,000 | SEC companyfacts, `Goodwill` |
| Finite-lived intangibles, net | 248,000,000 | 231,000,000 | SEC companyfacts, `FiniteLivedIntangibleAssetsNet` |
| Indefinite-lived intangibles, net | 26,641,000,000 | 27,415,000,000 | SEC companyfacts, `IndefiniteLivedIntangibleAssetsExcludingGoodwill` |
| Goodwill and intangibles total | 45,830,000,000 | 47,525,000,000 | Derived = goodwill + finite-lived + indefinite-lived intangibles |
| Operating income (EBIT proxy) | 23,595,000,000 | 23,994,000,000 | SEC companyfacts, `OperatingIncomeLoss` |
| Interest expense | 641,000,000 | 589,000,000 | SEC companyfacts, `InterestExpenseNonoperating` |
| Dividends paid | 4,217,000,000 | 4,634,000,000 | SEC companyfacts, `PaymentsOfDividends` |
| Dividend per share | 0.52 | 0.59 | SEC companyfacts, `CommonStockDividendsPerShareCashPaid` |
| Diluted EPS | n/a in companyfacts | 10.20 | Visa Form 10-K 2025, Note 16 EPS table |

## Mastercard

| Metric | 2024 | 2025 | Source |
|---|---:|---:|---|
| Revenue | 28,167,000,000 | 32,791,000,000 | SEC companyfacts, `Revenues` |
| Net income | 12,874,000,000 | 14,968,000,000 | SEC companyfacts, `ProfitLoss` |
| Total assets | 48,081,000,000 | 54,157,000,000 | SEC companyfacts, `Assets` |
| Current assets | 19,724,000,000 | 23,558,000,000 | SEC companyfacts, `AssetsCurrent` |
| Cash and cash equivalents | 8,442,000,000 | 10,566,000,000 | SEC companyfacts, `CashAndCashEquivalentsAtCarryingValue` |
| Accounts receivable, net | 3,773,000,000 | 4,609,000,000 | SEC companyfacts, `AccountsReceivableNetCurrent` |
| Current liabilities | 19,220,000,000 | 22,762,000,000 | SEC companyfacts, `LiabilitiesCurrent` |
| Total liabilities | 41,566,000,000 | 46,411,000,000 | SEC companyfacts, `Liabilities` |
| Total equity | 6,485,000,000 | 7,737,000,000 | SEC companyfacts, `StockholdersEquity` |
| Long-term debt current portion | 750,000,000 | 749,000,000 | SEC companyfacts, `LongTermDebtCurrent` |
| Long-term debt noncurrent | 17,476,000,000 | 18,251,000,000 | SEC companyfacts, `LongTermDebtNoncurrent` |
| Total debt used for project | 18,226,000,000 | 19,000,000,000 | Derived = current portion + noncurrent |
| PPE, net | 2,138,000,000 | 2,303,000,000 | SEC companyfacts, `PropertyPlantAndEquipmentNet` |
| Goodwill | 9,193,000,000 | 9,560,000,000 | SEC companyfacts, `Goodwill` |
| Finite-lived intangibles, net | 5,300,000,000 | 5,382,000,000 | SEC companyfacts, `FiniteLivedIntangibleAssetsNet` |
| Goodwill and intangibles total | 14,493,000,000 | 14,942,000,000 | Derived = goodwill + finite-lived intangibles |
| Operating income (EBIT proxy) | 15,582,000,000 | 18,897,000,000 | SEC companyfacts, `OperatingIncomeLoss` |
| Interest expense | 646,000,000 | 722,000,000 | SEC companyfacts, `InterestExpenseNonoperating` |
| Dividends paid | 2,448,000,000 | 2,756,000,000 | SEC companyfacts, `PaymentsOfDividends` |
| Dividend per share | 2.74 | 3.15 | SEC companyfacts, `CommonStockDividendsPerShareDeclared` |
| Diluted EPS | 13.89 | 16.52 | SEC companyfacts, `EarningsPerShareDiluted` |

## Points de methode deja fixes

- Pour les ratios melangeant bilan et compte de resultat, utiliser les moyennes `2024-2025`
- Pour `Total Debt / Total Assets`, la dette retenue = `LongTermDebtCurrent + LongTermDebtNoncurrent`
- Pour `Quick Ratio`, utiliser `Cash and cash equivalents + Accounts receivable`
- Pour `Interest coverage`, utiliser `Operating income / Interest expense`
- Pour `Goodwill and Intangibles`, additionner goodwill et actifs incorporels nets pertinents

## Points a traiter ensuite

- Verification de l'absence d'inventory chez Visa et Mastercard
  - aucune ligne inventory exploitable n'apparait dans les companyfacts SEC
  - le ratio `Inventory Turnover` est donc traite comme `n/a` pour ces deux groupes
