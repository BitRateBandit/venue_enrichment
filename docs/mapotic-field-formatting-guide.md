# Mapotic Place Import Field Formatting Guide

Date: 2026-04-06

This guide consolidates the Mapotic Help Center, public API docs, sample import sheets, and the public Discover experience into a production-oriented reference for preparing place data before import into Mapotic.

Use this guide with one key distinction:

- `Documented rule`: explicitly stated in Mapotic Help/API docs.
- `Safe convention`: not strictly documented by Mapotic, but the most reliable structure for imports and downstream rendering.

## Quick Reference

| Field | Type | Format | Example |
| --- | --- | --- | --- |
| Name | String | Free-text place name; required | `The Fillmore` |
| Address | String | Single address string for geocoding: `Road/Town/Postal Code` | `1805 Geary Blvd, San Francisco, CA 94115` |
| Latitude / Longitude | Number / Number | WGS84 (EPSG:4326), decimal degrees, separate columns | `37.7840` / `-122.4331` |
| Description | String | Long text / multiline text | `Historic music venue...` |
| Hours of Operation | String | Human-readable text; multiline recommended | `Mon-Fri 10:00-18:00` |
| Instagram | URL | Absolute `https://` profile or post URL | `https://www.instagram.com/thefillmore/` |
| Facebook | URL | Absolute `https://` URL | `https://www.facebook.com/TheFillmoreSF` |
| X (Twitter) | URL | Absolute `https://` URL | `https://x.com/thefillmore` |
| Email | Email string | Valid email format | `info@example.com` |
| YouTube | URL or Video ID object | Safe import: full YouTube URL; API video value: YouTube ID object | `https://www.youtube.com/watch?v=wkcglk95OzM` |
| Phone | String | Clickable phone number; international format recommended | `+1 415 555 0123` |
| Website | URL | Absolute `https://` URL | `https://www.thefillmore.com/` |
| TikTok | URL | Absolute `https://` URL | `https://www.tiktok.com/@thefillmore` |
| PlacesID | String or integer | Stable unique ID; Mapotic update syntax supports `mapotic:<id>` | `venue-000123` or `mapotic:00001` |
| City | String | Plain city name | `San Francisco` |
| State | String | Plain state name or abbreviation; stay consistent | `CA` |
| Zip Code | String | Postal code as text; preserve leading zeros | `94115` |
| Tickets Link | URL | Absolute `https://` URL | `https://www.ticketmaster.com/event/1A005F01ABCD1234` |
| Spotify URL | URL | Absolute `https://open.spotify.com/...` URL | `https://open.spotify.com/artist/1dfeR4HaWDbWqFHLkxsg1d` |
| Main Image URL | URL | One publicly accessible image URL | `https://cdn.example.com/main.jpg` |
| Additional Image URL(s) | URL list | Public image URLs separated with `$` | `https://cdn.example.com/1.jpg$https://cdn.example.com/2.jpg` |

## General Import Rules

- Supported import files are `XLS`, `CSV`, and `KML` per Mapotic Help.
- `Name` is mandatory.
- Location must be provided by either:
  - `Latitude` + `Longitude`, or
  - one full `Address` field.
- Custom fields only import cleanly if the matching attribute already exists on the target map.
- For clickable links in text fields, Mapotic documents that the value must include `http://`; in production, prefer `https://`.
- For images, the URL must be publicly accessible without login, cookies, or signed-session access.

> Warning: Mapotic does not publish hard character limits for most fields. Where no limit is documented below, treat the field as "no public Mapotic limit published" rather than assuming unlimited storage.

## Name

- Data type: `String`
- Required format / pattern: Free-text place name. This is a Mapotic mandatory field.
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `The Fillmore`, `Amoeba Music Hollywood`
- Notes or warnings: Keep this to the place name only. Do not overload it with city, category, or address details unless that is the actual public-facing name.

## Address

- Data type: `String`
- Required format / pattern: For geolocation import, Mapotic documents one combined address field in the format `Road/Town/Postal Code`.
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `1805 Geary Blvd, San Francisco, CA 94115`
- Notes or warnings: Mapotic states address-based geolocation is less accurate than GPS coordinates and is sensitive to misspellings.

> Warning: `City`, `State`, and `Zip Code` do not replace the geolocation `Address` field by themselves. If you want Mapotic to geocode from address text, pre-compose one full address string before import.

## Latitude / Longitude

- Data type: `Number` / `Number`
- Required format / pattern: WGS84 (`EPSG:4326`) in decimal degrees (`DD`), with latitude and longitude imported as separate fields.
- Character limits or constraints: Numeric values only.
- Accepted values or examples: `37.7840` and `-122.4331`
- Notes or warnings: Negative numbers indicate south/west. Do not use DMS (`37°47'02"N`) or projected coordinate systems.

> Warning: Mapotic import examples use separate `Latitude` and `Longitude` columns, but the public GeoJSON API returns coordinates in `[longitude, latitude]` order. Do not swap them during import preparation.

## Description

- Data type: `String`
- Required format / pattern: Long text / multiline text (`textarea`-style content).

### Content Standard (Required)

All descriptions must be written in the voice of a **Senior Rolling Stone Magazine Travel Section writer**.

This is not optional. The goal is to produce descriptions that feel editorial, immersive, and culturally grounded — not generic directory listings.

### Writing Requirements

- Must include exactly one mention of: `musicroadtrip.com`
- Must feel specific to the location (no generic filler text)
- Must include at least one experiential, sensory, or cultural detail
- Must read as human-written editorial, not templated or generated

### Structure & Style Rules

- Paragraph length: **60–120 words**
- Writing style must vary across entries:
  - Do not reuse opening phrases or sentence structures
  - Rotate narrative approach:
    - Scene-setting
    - Historical framing
    - Cultural significance
    - First-impression tone
- Avoid predictable openings such as:
  - "Nestled in the heart of"
  - "Located in"
  - "Known for"

### Prohibited Content Patterns

- Repetitive phrasing across records
- Generic tourism language
- SEO-style keyword stuffing
- Descriptions that could apply to multiple locations with minimal changes

### Quality Validation Rules (Enforced in Pipeline)

A description should be flagged if:

- It does NOT contain `musicroadtrip.com`
- It reuses a known banned phrase
- It is structurally similar to previously generated descriptions
- It lacks specificity (i.e., could describe another venue)

### Character Limits or Constraints

- No public Mapotic max length published
- Recommended operational range: **60–120 words**

### Accepted Example (Style Reference Only)

A dimly lit room where decades of feedback, distortion, and late-night sets seem baked into the walls, this venue carries the kind of lived-in authenticity you don’t manufacture—you inherit. Artists pass through, but the energy lingers, echoing in worn floors and patched-up amps. It’s the kind of place musicroadtrip.com exists to document, where the story matters as much as the sound, and every night feels like it could tip into something unforgettable.

### Notes or Warnings

- Mapotic renders this as long-form text; formatting such as line breaks is supported
- Embedded links must include full URLs (`https://`) to be clickable
- Treat this field as **editorial content**, not metadata

## Hours of Operation

- Data type: `String`
- Required format / pattern: Safe convention is human-readable text in a `Long text` attribute; Mapotic does not publish a dedicated opening-hours import schema.
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `Mon-Thu 11:00-22:00\nFri-Sat 11:00-00:00\nSun 11:00-21:00`
- Notes or warnings: Use one consistent schedule style across the whole dataset.

> Warning: There is no documented Mapotic importer format for structured recurring business hours such as `RRULE`, `OpeningHoursSpecification`, or Google-style hours objects. Treat this as display text unless you have a custom downstream parser.

## Instagram

- Data type: `URL string`
- Required format / pattern: Full absolute URL, preferably `https://www.instagram.com/<handle>/`
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `https://www.instagram.com/thefillmore/`
- Notes or warnings: Use a full URL, not just `@handle`. Mapotic only documents clickable links for text values that include a URL schema.

## Facebook

- Data type: `URL string`
- Required format / pattern: Full absolute URL, preferably `https://www.facebook.com/<page>`
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `https://www.facebook.com/TheFillmoreSF`
- Notes or warnings: Use the canonical public page URL, not a share URL shortened by Facebook.

## X (Twitter)

- Data type: `URL string`
- Required format / pattern: Full absolute URL, preferably `https://x.com/<handle>`
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `https://x.com/thefillmore`
- Notes or warnings: Do not import bare handles such as `@thefillmore`. If your source still uses `twitter.com`, normalize consistently.

## Email

- Data type: `Email string`
- Required format / pattern: Valid email format. Mapotic documents that this field accepts only valid emails.
- Character limits or constraints: Must pass email validation; no public Mapotic length limit published.
- Accepted values or examples: `info@example.com`
- Notes or warnings: Use one mailbox per field. If you need multiple contacts, use multiple attributes or a descriptive text field instead of comma-joining addresses.

## YouTube

- Data type: `URL string` for safest CSV import; API-native video value is an object containing a YouTube ID
- Required format / pattern: Safe convention for CSV import is a full YouTube URL such as `https://www.youtube.com/watch?v=<video_id>`. In the public API, Mapotic's `video` attribute value is documented as `{"youtube": "wkcglk95OzM"}`.
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `https://www.youtube.com/watch?v=wkcglk95OzM`, API value `{"youtube": "wkcglk95OzM"}`
- Notes or warnings: Mapotic publishes the API storage format for `video` attributes, but does not publish CSV import syntax for that attribute type.

> Warning: Do not assume the CSV importer accepts raw JSON for `Video` attributes unless you have tested it against the exact target map. For unattended imports, a plain URL field is the safer documented choice.

## Phone

- Data type: `String`
- Required format / pattern: Clickable phone number. Mapotic does not publish a strict mask; safe convention is one normalized number in international form.
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `+1 415 555 0123`, `+14155550123`
- Notes or warnings: Avoid labels such as `Box Office:` in the value itself. Keep extensions separate if possible.

## Website

- Data type: `URL string`
- Required format / pattern: Full absolute URL including protocol, preferably `https://`
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `https://www.thefillmore.com/`
- Notes or warnings: Do not use bare domains like `thefillmore.com`. Use canonical destination URLs, not tracking redirects when possible.

## TikTok

- Data type: `URL string`
- Required format / pattern: Full absolute URL, preferably `https://www.tiktok.com/@<handle>`
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `https://www.tiktok.com/@thefillmore`
- Notes or warnings: Use the public profile URL. Do not import only the handle.

## PlacesID

- Data type: `String` or `Integer`
- Required format / pattern: Stable unique identifier used for future updates. Mapotic explicitly documents numeric IDs such as `1` to `n`, and also supports `mapotic:<MapoticID>` for updating records that already exist in Mapotic.
- Character limits or constraints: Must be unique per place within the dataset used for updates; no public max length published.
- Accepted values or examples: `1`, `12345`, `venue-000123`, `mapotic:00001`
- Notes or warnings: The important rule is stability. Reuse the same ID on every future import if you want updates instead of duplicate place creation.

> Warning: If you change `PlacesID` values between imports, Mapotic will treat the row as a new place instead of an update target.

## City

- Data type: `String`
- Required format / pattern: Plain city name in a single-line text attribute.
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `San Francisco`
- Notes or warnings: This is best treated as a custom text attribute. It is not a standalone documented geolocation field for import resolution.

## State

- Data type: `String`
- Required format / pattern: Plain state name or abbreviation in a single-line text attribute.
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `CA`, `California`
- Notes or warnings: Pick one convention across the dataset and keep it consistent. For U.S. data, USPS two-letter abbreviations are the cleanest machine-readable option.

## Zip Code

- Data type: `String`
- Required format / pattern: Postal code stored as text.
- Character limits or constraints: Preserve leading zeros; no public Mapotic max length published.
- Accepted values or examples: `94115`, `02108`, `94115-3412`
- Notes or warnings: Treat zip codes as strings, not numbers, so spreadsheet tools do not strip leading zeros.

## Tickets Link

- Data type: `URL string`
- Required format / pattern: Full absolute URL including protocol, preferably `https://`
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `https://www.ticketmaster.com/event/1A005F01ABCD1234`
- Notes or warnings: Use the final public purchase page, not a session-bound cart URL.

## Spotify URL

- Data type: `URL string`
- Required format / pattern: Full absolute Spotify URL, typically `https://open.spotify.com/...`
- Character limits or constraints: No public Mapotic max length published.
- Accepted values or examples: `https://open.spotify.com/artist/1dfeR4HaWDbWqFHLkxsg1d`
- Notes or warnings: Use canonical Spotify URLs. Do not import `spotify:` URIs unless you have separately verified that your target rendering supports them.

## Main Image URL

- Data type: `URL string`
- Required format / pattern: Exactly one publicly accessible direct image URL mapped to Mapotic's `Main image` special parameter.
- Character limits or constraints:
  - Must resolve directly to an image asset
  - Must not resolve to an HTML page, gallery page, social post, or file-sharing landing page
  - Must not require login, cookies, JavaScript execution, referrer validation, or signed/expiring access
  - Must be stable enough for repeatable imports; avoid temporary or session-bound URLs
- Accepted values or examples: `https://cdn.example.com/images/fillmore-main.jpg`
- Selection guidance:
  - Prefer a wide, high-quality venue image that immediately communicates the space, atmosphere, and live-music identity of the location
  - First preference: a strong wide shot of the venue in action during a live performance, where the venue itself is clearly visible and not visually dominated by a single performer
  - Second preference: a clean exterior establishing shot of the facade, entrance, marquee, or building
  - The image should be venue-centric, not artist-centric
  - Favor landscape orientation and images that still read clearly in cropped card, map, and mobile layouts
- Notes or warnings:
  - Use a direct asset URL only
  - The main image should be the single best editorial representation of the venue
  - Prefer images that look professional, current, and geographically truthful to the venue

### Reject the image if any of the following are true

- It is a flyer, poster, event graphic, ad creative, or promotional banner
- It contains visible writing, large text overlays, dates, lineups, pricing, or branding copy
- It contains a watermark, photographer mark, agency stamp, logo bug, or visible copyright overlay
- It is primarily a food or drink photo
- It is focused on cocktails, beer, plates, menus, table settings, or hospitality details instead of the venue
- It is a close-up artist portrait or performance shot where the venue is not clearly legible
- It is a collage, composite, screenshot, meme, or social-media-style graphic
- It is heavily filtered, low-resolution, blurry, stretched, or badly cropped
- It shows signage only, without enough venue context
- It is not clearly a real photo of the actual venue

> Warning: Mapotic requires a publicly accessible image link. Authenticated CDN URLs, expiring signed URLs, Google Drive share pages, Dropbox preview pages, and social-media page URLs are unreliable import targets and must not be used.

## Additional Image URL(s)

- Data type: `URL list`
- Required format / pattern: One or more publicly accessible direct image URLs separated by a dollar sign (`$`)
- Character limits or constraints:
  - No leading `$`
  - No trailing `$`
  - No empty entries between separators
  - Do not include `$` separators when only one image is present
- Accepted values or examples: `https://cdn.example.com/1.jpg$https://cdn.example.com/2.jpg$https://cdn.example.com/3.jpg`
- Selection guidance:
  - The first URL in this field should be the exact same URL used for `Main Image URL`, unless the import workflow intentionally maps `Main image` separately
  - Preferred gallery order for venues:
    1. Main image URL first
    2. Exterior facade / entrance / marquee / signage shot
    3. Interior crowd or room shot
    4. Stage-focused image showing venue layout
    5. Distinct architectural or spatial feature that helps a traveler understand the place
  - Each image should add unique editorial value
  - Build a gallery that helps a user understand the venue’s exterior, interior, scale, layout, and live-music atmosphere
  - Keep all selected images venue-centric rather than performer-centric
- Notes or warnings:
  - Every URL must resolve directly and publicly to an image asset
  - Do not mix direct image URLs with page URLs
  - Avoid near-duplicates, repeated angles, and redundant crops
  - If the same image is used in both fields, the URL must match exactly

### Reject any gallery image if any of the following are true

- It is a flyer, poster, event card, ticket graphic, sponsor graphic, or promo asset
- It contains visible writing, captions, lineups, dates, price text, or text-heavy signage as the main focus
- It contains watermarks, logos, or visible ownership overlays
- It is primarily food, drinks, bottle service, menus, or tabletop content
- It is primarily an artist press photo rather than a venue photo
- It is a duplicate or near-duplicate of another selected image
- It is too dark, blurry, noisy, pixelated, distorted, or cropped so tightly that venue context is lost
- It is not clearly tied to the real venue
- It adds no new information about the venue’s appearance, atmosphere, or function

> Warning: Mapotic documents that the first image in the list may become the main image unless `Main image` is mapped separately. To remove an image gallery via import, use the sentinel value `$DELETE` by itself. Never combine `$DELETE` with actual image URLs in the same field.

## Implementation Notes for Import Pipelines

- Prefer GPS coordinates over address geocoding whenever exact placement matters.
- Create all required Mapotic custom attributes before import. Import mapping cannot target attributes that do not exist.
- Normalize all social, ticketing, and music links to full canonical `https://` URLs before writing the import file.
- Preserve `Zip Code` as text in spreadsheet or CSV generation code.
- Keep `PlacesID` stable forever once assigned.
- For image imports, preflight URLs with an unauthenticated HTTP `200` check and a valid image content type.
- Treat `YouTube` as a plain URL field unless you have positively validated Mapotic `video` attribute CSV behavior in the target environment.

## Sources Reviewed

- [Mapotic Help: How to bulk import places into your Mapotic map](https://help.mapotic.com/import-data-places-mapotic-map/)
- [Mapotic Help: Attribute Types](https://help.mapotic.com/attribute-types/)
- [Mapotic Help: Setting up categories and attributes](https://help.mapotic.com/setting-up-categories-attributes/)
- [Mapotic API: Attribute](https://mapotic.github.io/mapotic.com-api-docs/attribute/)
- [Mapotic API: POI](https://mapotic.github.io/mapotic.com-api-docs/poi/)
- [Mapotic Discover](https://www.mapotic.com/discover)

## Confidence Notes

- `High confidence`: `Name`, `Address`, `Latitude / Longitude`, `PlacesID`, image fields, email validation, API attribute value types.
- `Medium confidence`: social URLs, `Website`, `Tickets Link`, `Spotify URL`, `Hours of Operation`, `City`, `State`, `Zip Code`, because Mapotic does not publish stricter field-specific validation beyond generic text/link behavior.
- `Caution required`: `YouTube`, because Mapotic publishes the API storage shape for `video` attributes but not the CSV import syntax for that attribute type.
