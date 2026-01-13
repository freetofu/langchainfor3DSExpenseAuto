# Expense mapping configuration

This file documents how `expenseMapping.json` is used by `popupScript.js` when a
raw sales statement Excel file is loaded.

## Defaults

`defaults` provide fallback values used when the raw file does not include a
field that is required by the expense input form.

- `paymentType`: Default payment type to send to the form.
- `billingType`: Default billing type to send to the form.
- `currency`: Default currency when `통화코드` is missing.
- `expenseTypeFallback`: Used when no rule matches a row.

### Fixed defaults applied
- `paymentType`: `AMEX/CASH`
- `billingType`: `CorpNonBillable`
- `currency`: `KRW`
- `expenseTypeFallback`: `Miscellaneous`

## Rules

Each rule can match on merchant name, usage place, category name/code, user name,
amount range, and approval time.

```json
{
  "expenseType": "Meals - hosp.",
  "match": {
    "merchantName": ["스타벅스"],
    "usagePlace": ["카페"],
    "categoryName": ["음식"],
    "categoryCode": ["1000"],
    "userName": ["홍길동"]
  },
  "amountRange": { "min": 10000, "max": 50000 },
  "timeRange": { "start": "18:00", "end": "23:59" }
}
```

- The first matching rule wins.
- Omit `amountRange`/`timeRange` if they are not required for a rule.
- Any expense type containing `hosp.` will be highlighted in the preview so the
  user can enter attendee details.

### Time-range assumptions used in current mapping
- Breakfast: 06:00–10:59
- Lunch: 11:00–14:59
- Dinner: 18:00–22:59

### Special handling
- Taxi rows that occur within **10 seconds** of each other on the same date and
  same user are merged into a single item by summing the amounts.
