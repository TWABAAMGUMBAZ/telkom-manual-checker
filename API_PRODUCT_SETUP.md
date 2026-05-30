# Telkom Checker API Setup

Base URL:

```text
https://telkom-manual-checker.onrender.com
```

Authentication:

```http
X-API-Key: Luna2023
```

You can also use:

```http
Authorization: Bearer Luna2023
```

For production, set a separate `API_KEYS` environment variable in Render with one or more comma-separated customer keys.

## Single Number Lookup

```bash
curl -X POST "https://telkom-manual-checker.onrender.com/v1/lookup" \
  -H "X-API-Key: Luna2023" \
  -H "Content-Type: application/json" \
  -d "{\"number\":\"+27 12 791 3714\"}"
```

Response:

```json
{
  "input_number": "+27 12 791 3714",
  "clean_number": "0127913714",
  "number_type": "Tshwane 012",
  "lookup_status": "Found",
  "current_provider": "TELKOM",
  "telkom_service": true,
  "raw_result": "Number query result ...",
  "checked_at": "2026-05-30 08:00:00",
  "porting_lookup_url": "https://www.porting.co.za/PublicWebsiteApp/#/number-inquiry?sid=smppipd4x1&msisdn=0127913714"
}
```

## Batch Job

```bash
curl -X POST "https://telkom-manual-checker.onrender.com/v1/batch" \
  -H "X-API-Key: Luna2023" \
  -H "Content-Type: application/json" \
  -d "{\"numbers\":[\"+27 11 463 6368\",\"+27 12 791 3714\",\"+27 82 123 4567\"]}"
```

Returns immediately:

```json
{
  "job_id": "abc123",
  "status": "queued",
  "total": 3,
  "processed": 0,
  "message": "Queued."
}
```

Check progress:

```bash
curl "https://telkom-manual-checker.onrender.com/v1/jobs/abc123" \
  -H "X-API-Key: Luna2023"
```

Get results:

```bash
curl "https://telkom-manual-checker.onrender.com/v1/jobs/abc123/results" \
  -H "X-API-Key: Luna2023"
```

## File Upload API

```bash
curl -X POST "https://telkom-manual-checker.onrender.com/v1/files" \
  -H "X-API-Key: Luna2023" \
  -F "file=@numbers.xlsx"
```

The API reads the first sheet for Excel, or header row for CSV, extracts supported `011`, `012`, and South African mobile numbers, then starts a background job.

## Status Values

- `Found`: provider was returned.
- `unsupported_number`: number is not `011`, `012`, or supported mobile.
- `needs_human_verification`: the Porting service requested captcha/manual verification.
- `needs_manual_review`: the lookup returned an unclear response.

## Production Notes

This API is now job-based and responsive, but the current lookup engine still depends on the public Porting flow. For a commercial resale API, replace the lookup engine with a licensed South African operator/porting data provider. The customer-facing endpoints can stay the same.
