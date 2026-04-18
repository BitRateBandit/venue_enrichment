# JamBase API Reference (Live-Verified)

Generated from live API calls on **2026-04-10** using `JAMBASE_API_KEY`.  
Source artifacts are in `/Users/augat/codex-jambase/data/debug/`.

## 1. Overview

The live API currently returns:

- Event records (`Concert`, `Festival`)
- Venue records
- Artist records
- Stream records

From the sampled `events` responses, coverage is clearly international (examples included `AT`, `GB`, `DE`, `BR`, `NZ`, `MX`, `FR`, `ES`, `IT`, plus `US`).

## 2. Base Configuration

- Base URL (verified): `https://www.jambase.com/jb-api/v1`
- Authentication (verified): `apikey` query parameter is required
- Header requirement (verified): calls succeed with an explicit `User-Agent`

Observed request envelope:

- `request.endpoint`
- `request.method`
- `request.params`
- `request.userAgent`

## 3. Verified Endpoints

### Working

- `GET /events`
- `GET /events/id/{id}`
- `GET /artists` (when required search params are present)
- `GET /artists/id/{id}`
- `GET /venues` (when required search params are present)
- `GET /venues/id/{id}`
- `GET /streams`

### Partial / Constraints

- `GET /artists` without any of `artistId`, `artistName`, or `genreSlug` returns `400 parameter_missing`.
- `GET /venues` without `venueName`, `venueId`, or geo params returns `400 parameter_missing`.

### Not Supported (from this run)

- No additional endpoints were confirmed beyond the list above.

## 4. Query Parameters

### Confirmed Working

- Global auth:
  - `apikey`
- Events:
  - `page`
  - `perPage`
  - `geoCountryIso2`
  - `eventType=concert`
  - `eventType=festival`
  - `eventDatePreset=today`
- Artists:
  - `artistId`
- Venues:
  - `geoCountryIso2`

### Partially Working

- `eventDatePreset=upcoming` was rejected with:
  - `400 invalid_param`
  - message: `The date preset 'upcoming' is not valid.`

### Mentioned by API Errors (not fully tested in this run)

- Artists endpoint error mentions: `artistName`, `genreSlug`
- Venues endpoint error mentions: `venueName`, `venueId`, `geoCityId`, `geoCountryIso3`, `geoIp`, `geoMetroId`, `geoStateIso`, `geoLatitude` + `geoLongitude`

## 5. Pagination Model

Verified on `/events`:

- Response includes `pagination` with:
  - `page`
  - `perPage`
  - `totalItems`
  - `totalPages`
  - `nextPage` (absolute URL)
  - `previousPage`
- `page=2` works.
- `nextPage` URL works directly.
- `perPage` supports at least `1..100`.
- `perPage=200` returns `400 invalid_param` (`Please use a number between 1 and 100.`).

Recommended approach:

1. Start with `/events?apikey=...&page=1&perPage=100`.
2. Continue while events are returned and `pagination.nextPage` is present.
3. Prefer explicit `page` increments for deterministic replay/checkpointing.

## 6. Full Event Schema

### Top-level event fields (observed)

- `@type` (`Concert` or `Festival` in samples)
- `identifier` (e.g., `jambase:15103401`)
- `name`
- `startDate`
- `endDate`
- `eventStatus`
- `doorTime`
- `eventAttendanceMode`
- `url`
- `image`
- `isAccessibleForFree`
- `datePublished`
- `dateModified`
- `previousStartDate`
- `location` (object)
- `performer` (array)
- `offers` (array)
- `sameAs` (array)
- `x-*` fields observed:
  - `x-customTitle`
  - `x-headlinerInSupport`
  - `x-lineupDisplayOption`
  - `x-promoImage`
  - `x-streamIds`
  - `x-subtitle`

### Location / venue sub-structure (observed)

- `location.identifier`
- `location.name`
- `location.url`
- `location.image`
- `location.maximumAttendeeCapacity`
- `location.geo.latitude`
- `location.geo.longitude`
- `location.address.streetAddress`
- `location.address.x-streetAddress2`
- `location.address.addressLocality`
- `location.address.addressRegion` (object, sometimes empty `{}`)
- `location.address.postalCode`
- `location.address.addressCountry.identifier` (ISO2)
- `location.address.addressCountry.alternateName` (ISO3-like)
- `location.address.addressCountry.name`
- `location.address.x-timezone`
- `location.sameAs` (array)
- `location.x-isPermanentlyClosed`
- `location.x-numUpcomingEvents`

### Performer structure (observed)

- `performer[].@type`
- `performer[].identifier`
- `performer[].name`
- `performer[].url`
- `performer[].image`
- `performer[].genre` (array)
- `performer[].x-bandOrMusician`
- `performer[].x-isHeadliner`
- `performer[].x-performanceDate`
- `performer[].x-performanceRank`
- `performer[].x-dateIsConfirmed`
- `performer[].x-numUpcomingEvents`

### Offers structure (observed)

- `offers[].@type`
- `offers[].identifier`
- `offers[].name`
- `offers[].category` (e.g., `ticketingLinkPrimary`)
- `offers[].url`
- `offers[].validFrom`
- `offers[].priceSpecification.price`
- `offers[].priceSpecification.minPrice`
- `offers[].priceSpecification.maxPrice`
- `offers[].priceSpecification.priceCurrency`
- `offers[].seller.identifier`
- `offers[].seller.name`
- `offers[].seller.disambiguatingDescription`
- `offers[].x-spansDays`

### Socials (`sameAs`) structure (observed)

- `sameAs[].identifier` examples:
  - `officialSite`
  - `facebook`
  - `instagram`
  - `spotify`
- `sameAs[].url`

### Real examples (trimmed)

```json
{
  "@type": "Festival",
  "identifier": "jambase:15103401",
  "name": "Snowbombing",
  "startDate": "2026-04-06",
  "endDate": "2026-04-11",
  "eventStatus": "scheduled",
  "location": {
    "name": "Brück’n Stadl",
    "address": {
      "addressCountry": {"identifier": "AT", "name": "Austria"},
      "addressLocality": "Mayrhofen",
      "postalCode": "6290"
    }
  }
}
```

## 7. Festival vs Concert Logic

Live responses showed both mechanisms working:

- `@type` on event records (`Concert` / `Festival`)
- `eventType` query filter accepted:
  - `eventType=concert`
  - `eventType=festival`

Recommended detection:

1. Primary: trust `event.@type`.
2. Query optimization: use `eventType` when intentionally narrowing results.

## 8. Non-US Filtering Strategy

Reliable approach from live schema:

1. Prefer API pre-filter when possible:
   - `geoCountryIso2=<ISO2>`
2. Always apply post-filter guard using `location.address.addressCountry`:
   - reject if any normalized value equals:
     - `US`
     - `USA`
     - `UNITED STATES`
     - `UNITED STATES OF AMERICA`
   - normalize/check against:
     - `addressCountry.identifier`
     - `addressCountry.alternateName`
     - `addressCountry.name`

## 9. Known Limitations

- `eventDatePreset=upcoming` is invalid (at least in this environment/date).
- `x-jb-api-requests-remaining` was not present in sampled response headers.
- Some nested objects can be empty (`addressRegion: {}`).
- Optional event fields vary (`x-lineupDisplayOption` appeared in only part of sampled events).
- `sameAs`, `offers`, and `performer` can be empty arrays.
- Search endpoints require endpoint-specific params and return `400` if missing.

## 10. Best Practices

- Always include:
  - `apikey` query param
  - explicit `User-Agent` header
- Preserve raw JSON (`jsonl`) for traceability.
- Parse defensively:
  - treat all nested fields as optional
  - allow empty objects/arrays
- Deduplicate by `event.identifier`.
- Use checkpointing with `page` and parameter segment context.
- Apply non-US exclusion before output write.

## 11. Integration Notes (Mapotic / CSV / Flattening)

- Keep one raw JSON copy per event row.
- Normalize URLs to absolute `https://` where possible.
- Preserve postal codes as text.
- Keep latitude/longitude in decimal degrees; never swap order.
- Join multi-value image URLs with `$`.
- Keep both:
  - normalized operational columns
  - wide flattened columns (`location.address.streetAddress`, `offers[0].seller.name`, etc.).

---

## Verification Artifacts

- `/Users/augat/codex-jambase/data/debug/events_sample_page1.json`
- `/Users/augat/codex-jambase/data/debug/pagination_recheck.json`
- `/Users/augat/codex-jambase/data/debug/parameter_tests.json`
- `/Users/augat/codex-jambase/data/debug/endpoint_tests.json`
- `/Users/augat/codex-jambase/data/debug/endpoint_payload_index.json`
- `/Users/augat/codex-jambase/data/debug/events_id.json`
- `/Users/augat/codex-jambase/data/debug/venues_search_geoCountryIso2_GB.json`
- `/Users/augat/codex-jambase/data/debug/venues_id.json`
- `/Users/augat/codex-jambase/data/debug/artists_search_artistId.json`
- `/Users/augat/codex-jambase/data/debug/artists_id.json`
- `/Users/augat/codex-jambase/data/debug/streams.json`
- `/Users/augat/codex-jambase/data/debug/schema_samples.json`
- `/Users/augat/codex-jambase/data/debug/live_discovery_summary.json`
