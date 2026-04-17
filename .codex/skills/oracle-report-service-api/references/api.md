# API Reference

Base workflow:

1. `GET /api/health`
2. `GET /api/reports/sheets`
3. `POST /api/reports/generate`
4. `GET /api/reports/download/:fileId`
5. `GET /api/reports/debug/unmatched`

## Generate

```http
POST /api/reports/generate
Content-Type: application/json
Authorization: Bearer <token>
```

Body:

```json
{
  "month": "2025-12",
  "sheetName": "排名",
  "sheetOnly": true
}
```

## Download

Use the `file.id` returned from generate:

```http
GET /api/reports/download/:fileId
Authorization: Bearer <token>
```

## Debug unmatched

```http
GET /api/reports/debug/unmatched?month=2025-12&limit=20
Authorization: Bearer <token>
```
