# Interview booking — SharePoint + Power Automate setup

This guide connects `interview-booking.html` to two SharePoint lists and two Power Automate flows. The page is deployed on the same Azure Static Web App as the internship form.

**Booking page URL (after deploy):**

```
https://<your-static-web-app>.azurestaticapps.net/interview-booking.html?email=intern@example.com
```

Use `encodeURIComponent(email)` when building links in Power Automate or Outlook.

---

## 1. SharePoint lists

Create both lists on the same SharePoint site you use for HR/internship data.

### List A: `Interview Slots`

| Column display name | Type | Internal name | Notes |
|---------------------|------|---------------|-------|
| Title | Single line of text | `Title` | e.g. "Interview slot 1" |
| Start Date Time | Date and Time | `StartDateTime` | Date & time, **UTC** or site timezone (be consistent) |
| End Date Time | Date and Time | `EndDateTime` | Optional |
| Max Bookings | Number | `MaxBookings` | Use `1` for one intern per slot (recommended for 24 interns) |
| Is Active | Yes/No | `IsActive` | Only `Yes` rows appear on the page |
| Display Order | Number | `DisplayOrder` | Sort order (1, 2, 3…) |
| Time Zone | Single line of text | `TimeZone` | Optional, e.g. `Europe/Belgrade` (shown on the page) |

**Example:** 24 rows — one row per 30‑minute interview, `MaxBookings = 1`, `IsActive = Yes`.

### List B: `Interview Bookings`

| Column display name | Type | Internal name | Notes |
|---------------------|------|---------------|-------|
| Title | Single line of text | `Title` | Set to candidate email |
| Candidate Email | Single line of text | `CandidateEmail` | Must match URL parameter |
| Slot ID | Number | `SlotId` | ID of the item in **Interview Slots** |
| Slot Start | Date and Time | `SlotStart` | Copy from slot at booking time |
| Booked At | Date and Time | `BookedAt` | `utcNow()` when flow runs |
| Status | Choice | `Status` | Choices: `Confirmed`, `Cancelled` (default `Confirmed`) |

---

## 2. Power Automate flow — List slots

**Name:** `Interview Booking - List Slots`

1. **Trigger:** When an HTTP request is received  
   - Method: **POST**  
   - Request body JSON schema:

```json
{
  "type": "object",
  "properties": {
    "email": { "type": "string" }
  }
}
```

2. **Initialize variable** `NormalizedEmail` = `toLower(trim(triggerBody()?['email']))`

3. **Get items** — SharePoint `Interview Bookings`  
   - Filter: `CandidateEmail eq '{NormalizedEmail}' and Status eq 'Confirmed'`  
   - Top: 1  

4. **Condition:** length of bookings > 0  

   **If yes** — respond with existing booking:

```json
{
  "booking": {
    "slotId": @{items('Get_bookings')?['SlotId']},
    "title": "...",
    "label": "...",
    "startDateTime": "@{items('Get_bookings')?['SlotStart']}",
    "timeZone": "Europe/Belgrade"
  },
  "slots": []
}
```

   Load slot title from **Interview Slots** by `SlotId` for `title` / `label`.

   **If no** — continue below.

5. **Get items** — SharePoint `Interview Slots`  
   - Filter: `IsActive eq true`  
   - Order by: `DisplayOrder` ascending  

6. **For each** slot (Apply to each):

   - **Get items** — `Interview Bookings`  
     - Filter: `SlotId eq '{current Slot ID}' and Status eq 'Confirmed'`  
   - **Compose** `bookingCount` = length of that result  
   - **Compose** `remaining` = `MaxBookings - bookingCount`  
   - **Append to array variable** `SlotResults` (initialize before loop):

```json
{
  "id": @{items('Apply_to_each')?['ID']},
  "title": "@{items('Apply_to_each')?['Title']}",
  "startDateTime": "@{items('Apply_to_each')?['StartDateTime']}",
  "endDateTime": "@{items('Apply_to_each')?['EndDateTime']}",
  "timeZone": "@{coalesce(items('Apply_to_each')?['TimeZone'], 'Europe/Belgrade')}",
  "remaining": @{variables('remaining')},
  "available": @{greater(variables('remaining'), 0)}
}
```

7. **Response** — HTTP 200, `application/json`:

```json
{
  "slots": @{variables('SlotResults')},
  "booking": null
}
```

Copy the **HTTP POST URL** into `interview-booking.js` → `listSlotsFlowUrl`.

---

## 3. Power Automate flow — Book slot

**Name:** `Interview Booking - Book Slot`

1. **Trigger:** When an HTTP request is received  
   - Method: **POST**  
   - Schema:

```json
{
  "type": "object",
  "properties": {
    "email": { "type": "string" },
    "slotId": { "type": "integer" }
  },
  "required": ["email", "slotId"]
}
```

2. **Normalize email** (same as above).

3. **Get items** — `Interview Bookings` where `CandidateEmail eq email` and `Status eq 'Confirmed'`  
   - If any → **Response** 409:

```json
{ "error": "already_booked", "message": "You already have an interview scheduled." }
```

4. **Get item** — `Interview Slots` by `slotId` from body.  
   - If not found or `IsActive` is false → **Response** 404.

5. **Get items** — bookings for this `SlotId` with `Status eq 'Confirmed'`.  
   - If `count >= MaxBookings` → **Response** 409:

```json
{ "error": "slot_full", "message": "This time slot is no longer available." }
```

6. **Create item** — `Interview Bookings`:

| Field | Value |
|-------|--------|
| Title | normalized email |
| CandidateEmail | normalized email |
| SlotId | slotId from body |
| SlotStart | slot's StartDateTime |
| BookedAt | `utcNow()` |
| Status | `Confirmed` |

7. **Response** 200:

```json
{
  "success": true,
  "booking": {
    "slotId": @{body('Create_item')?['SlotId']},
    "title": "@{body('Get_slot')?['Title']}",
    "label": "@{body('Get_slot')?['Title']}",
    "startDateTime": "@{body('Get_slot')?['StartDateTime']}",
    "timeZone": "@{coalesce(body('Get_slot')?['TimeZone'], 'Europe/Belgrade')}"
  }
}
```

Copy the **HTTP POST URL** into `interview-booking.js` → `bookSlotFlowUrl`.

---

## 4. Configure the web page

Edit `interview-booking.js`:

```javascript
const INTERVIEW_BOOKING_CONFIG = {
  listSlotsFlowUrl: "https://...powerplatform.com/.../invoke?...",
  bookSlotFlowUrl: "https://...powerplatform.com/.../invoke?..."
};
```

Commit and push to `main` — Azure Static Web Apps deploys automatically.

---

## 5. Email invitation (Power Automate or Outlook)

**Link pattern:**

```
https://<your-app>.azurestaticapps.net/interview-booking.html?email=@{encodeUriComponent(outputs('CandidateEmail'))}
```

**Example body:**

> Please select your interview time using your personal link (do not share this link):  
> [Choose interview time](https://.../interview-booking.html?email=...)

You can send 24 emails from:

- An **Instant cloud flow** with "Manually trigger" + Excel/SharePoint list of emails, or  
- A flow triggered when internship application status = "Invite to interview", or  
- Outlook mail merge / Power Automate **Send an email (V2)** in a loop.

---

## 6. Optional: confirmation email after booking

In the **Book slot** flow, after **Create item**, add:

- **Send an email (V2)** to `CandidateEmail` with date/time and Teams/meeting details  
- Or **Create event (V3)** in Outlook calendar  

---

## 7. Managing slots

| Task | Where |
|------|--------|
| Add / change interview times | Edit rows in **Interview Slots** |
| Disable a slot | Set **Is Active** = No |
| See who booked | **Interview Bookings** list |
| Free a slot (reschedule) | Set booking **Status** = `Cancelled`; intern can book again |

---

## 8. Security notes

- The link is **identified by email in the URL**, not a signed token. Do not publish the booking URL publicly; send only to invited interns.  
- Flows should **normalize and validate** email format.  
- Enforce **one Confirmed booking per email** in the book flow (steps above).  
- For stronger security later, add a one-time `token` column on invitations and validate it in both flows.

---

## 9. Testing locally

```bash
python3 -m http.server 8080
```

Open:

```
http://localhost:8080/interview-booking.html?email=test@example.com
```

Until flow URLs are configured, the page shows a configuration error — expected.

---

## 10. API contract (reference)

### POST List slots

**Request:** `{ "email": "intern@uni.edu" }`  

**Response (available):**

```json
{
  "slots": [
    {
      "id": 12,
      "title": "Morning slot",
      "startDateTime": "2026-06-15T10:00:00Z",
      "endDateTime": "2026-06-15T10:30:00Z",
      "timeZone": "Europe/Belgrade",
      "remaining": 1,
      "available": true
    }
  ],
  "booking": null
}
```

**Response (already booked):**

```json
{
  "slots": [],
  "booking": {
    "slotId": 12,
    "title": "Morning slot",
    "startDateTime": "2026-06-15T10:00:00Z",
    "timeZone": "Europe/Belgrade"
  }
}
```

### POST Book slot

**Request:** `{ "email": "intern@uni.edu", "slotId": 12 }`  

**Response:** `{ "success": true, "booking": { ... } }`

**Errors:** `409` with `{ "message": "..." }` for already booked or slot full.
