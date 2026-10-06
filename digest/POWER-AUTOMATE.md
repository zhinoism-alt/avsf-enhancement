# Daily AVSF digest with Power Automate

Sends the digest from your own Outlook mailbox (no SMTP login needed). About 15 minutes to build.

You need: Power Automate with the **HTTP** action (premium connector). The query below
was tested against the live Firestore project and returns your open requests.

> Works because the Firestore rules allow reads. If you later restrict the rules, this
> flow will start failing with HTTP 403; use the GitHub digest with a service account then.

## Build the flow

Power Automate → **Create** → **Scheduled cloud flow** → name it `AVSF daily digest`,
repeat every **1 week** on **Mon–Fri** at **07:00** (your time zone).

### 1. Initialize variable
- Name `StuckDays`, Type `Integer`, Value `7`

### 2. HTTP
- Method: `POST`
- URI (replace `YOUR_API_KEY` with the `apiKey` value near the top of the
  `firebaseConfig` block in `avsf_dashboard.html`):

```
https://firestore.googleapis.com/v1/projects/avsf-dashboard/databases/(default)/documents:runQuery?key=YOUR_API_KEY
```
- Headers: `Content-Type` = `application/json`
- Body:

```json
{
  "structuredQuery": {
    "from": [{ "collectionId": "avsfs" }],
    "where": {
      "fieldFilter": {
        "field": { "fieldPath": "status" },
        "op": "IN",
        "value": { "arrayValue": { "values": [
          { "stringValue": "PROCESSING" },
          { "stringValue": "ON_HOLD" },
          { "stringValue": "PENDING_BUYER" }
        ] } }
      }
    }
  }
}
```
(IDOC errors are left out on purpose: VM already completed them.)

### 3. Filter array — name it `Docs`
- From: `body('HTTP')`
- Condition (edit in advanced mode):
```
@not(empty(item()?['document']))
```
(Firestore adds one empty row at the end of the result.)

### 4. Select — name it `Rows`
- From: `body('Docs')`
- Switch the Map to **text mode** (the `T` icon) and paste:

```
{
  "Age": @{int(div(sub(ticks(utcNow('yyyy-MM-dd')), ticks(substring(coalesce(item()?['document']?['fields']?['entryDate']?['stringValue'], utcNow()), 0, 10))), 864000000000))},
  "Request": "@{item()?['document']?['fields']?['requestNo']?['stringValue']}",
  "Vendor": "@{item()?['document']?['fields']?['vendorName']?['stringValue']}",
  "CC": "@{item()?['document']?['fields']?['companyCode']?['stringValue']}",
  "Status": "@{item()?['document']?['fields']?['status']?['stringValue']}",
  "With": "@{coalesce(item()?['document']?['fields']?['approverName']?['stringValue'], '')}",
  "Rank": @{sub(100000, int(div(sub(ticks(utcNow('yyyy-MM-dd')), ticks(substring(coalesce(item()?['document']?['fields']?['entryDate']?['stringValue'], utcNow()), 0, 10))), 864000000000)))},
  "Waiting": @{or(or(equals(item()?['document']?['fields']?['status']?['stringValue'], 'ON_HOLD'), equals(item()?['document']?['fields']?['status']?['stringValue'], 'PENDING_BUYER')), and(not(empty(coalesce(item()?['document']?['fields']?['approverName']?['stringValue'], ''))), and(not(startsWith(toUpper(coalesce(item()?['document']?['fields']?['approverName']?['stringValue'], '')), 'VM TEAM')), and(not(contains(toUpper(coalesce(item()?['document']?['fields']?['approverName']?['stringValue'], '')), 'MUNOZ')), and(not(contains(toUpper(coalesce(item()?['document']?['fields']?['approverName']?['stringValue'], '')), 'RAMOS')), not(contains(toUpper(coalesce(item()?['document']?['fields']?['approverName']?['stringValue'], '')), 'ROJAS'))))))) }
}
```
`Rank` makes the oldest request sort first. `Waiting` is true for On hold, Pending buyer,
and Processing requests that SAP shows with someone outside the VM team
(same rule as the dashboard).

### 5. Filter array — name it `Stuck`
- From: `body('Rows')`
- Condition (advanced mode):
```
@greaterOrEquals(item()?['Age'], variables('StuckDays'))
```

### 6. Condition — `Anything stuck?`
- `length(body('Stuck'))` **is greater than** `0`
- Everything below goes in the **If yes** branch.

### 7. Filter array — `Ours` (inside If yes)
- From: `sort(body('Stuck'), 'Rank')`
- Condition: `@equals(item()?['Waiting'], false)`

### 8. Filter array — `Others` (inside If yes)
- From: `sort(body('Stuck'), 'Rank')`
- Condition: `@equals(item()?['Waiting'], true)`

### 9. Create HTML table — `Table ours`
- From: `body('Ours')`, Columns: **Custom**
  - `Age` → `@{concat(item()?['Age'], 'd')}`
  - `Request` → `@{item()?['Request']}`
  - `Vendor` → `@{item()?['Vendor']}`
  - `CC` → `@{item()?['CC']}`
  - `Status` → `@{item()?['Status']}`
  - `With` → `@{item()?['With']}`

### 10. Create HTML table — `Table others`
Same as step 9 but From `body('Others')`.

### 11. Send an email (V2) (Office 365 Outlook)
- To: your team's addresses (`;` separated)
- Subject:
```
AVSF digest: @{length(body('Stuck'))} request(s) open @{variables('StuckDays')}+ days
```
- Body (use the `</>` code view, then paste):
```html
<div style="font-family:Segoe UI,Arial,sans-serif;">
<h2>AVSF daily digest</h2>
<p>@{length(body('Stuck'))} request(s) have been open @{variables('StuckDays')}+ days (oldest first).</p>
<h3>In our queue (@{length(body('Ours'))})</h3>
@{body('Table ours')}
<h3>Waiting on others (@{length(body('Others'))})</h3>
@{body('Table others')}
<p><a href="YOUR_DASHBOARD_URL">Open the dashboard</a></p>
</div>
```

## Test it
Save, then **Test → Manually → Run flow**. With your current data you should get one
email listing the requests open 7+ days. Change `StuckDays` to change the threshold.

## Optional
- Teams instead of email: replace step 11 with **Post message in a chat or channel**,
  and use the same table output.
- Nicer table: wrap `body('Table ours')` in `replace(body('Table ours'), '<table>', '<table style="border-collapse:collapse;font-size:13px;">')`.
