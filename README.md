# CALGAS Capacitors — Production Tracker 3.2

Material moving through the factory, one department to the next, and back when it has to go back.

```
Purchase → Stores → Winding → Spray → Short Clearance → Assembly-MFD / Assembly-kVAr → Stores → Dispatch
```

Metallization and Slitting are gone. Nothing is written for them any more, and their old records stay in the sheet as history.

## What changed in 3.2

**Looks and works like Stock Management.** Same design, menu (collapsible on a computer, a bottom bar on a phone), loading screen, dropdowns that reopen on a click, checkbox filters, sortable tables with paging, and Excel export. On a phone, every table turns into cards, so Receive, Reject and Send are always on screen.

**Five pages:**

- **Overview** — what's waiting for you, what's on the move, open and late orders, and elements on hand. Four charts: elements along the route, orders by status, handovers received over the last 14 days, and open handovers by how long they've waited. **Needs attention** lists late orders, orders due this week, and handovers left waiting two days or more. Each one opens where it can be dealt with.
- **Orders** — the whole order book, with search and filters. Tap an order to see what has gone to Dispatch against it, and to edit, start, close or cancel it.
- **Handovers** — one department's work: what's waiting, elements on hand (a tile per spec, plus recent entries), what it has sent or asked for, and what's recently settled. **Send or return**, **Ask a neighbour** and **Record elements** open as forms.
- **History** — every handover and element entry, filterable by department, kind, status and dates, and exportable.
- **Trace** — type an order (ORD-…), a handover (HO-…), an item code or an element spec. You see where it went, who handled it and when, and for a spec, how many each department holds now. Old element codes (ELM-…) still show their history, for Managers and Admins.

**Orders move on.** An order goes Pending → In production → Part dispatched → Dispatched, and can be Closed (short-closed, with a reason) or Cancelled (with a reason, only while nothing has gone). "Dispatched" is counted from the Stores → Dispatch handovers that name the order, so a dispatch, a return or a correction moves it by itself. When Stores sends finished goods to Dispatch, the form offers the open orders for that item. A wrong item, an order that's closed, or more than is left is refused. Orders can be edited (customer, quantity, due date, the customer's PO, remarks), but never below what has been dispatched. Orders past their due date are marked **Late**.

**Part receipts.** The receiver can take less than was sent, with a note saying why (e.g. 95 of 100, "5 damaged"). Stock Management is posted only for what arrived. The rest becomes its own handover, still on its way from the sender, so it can be received later or rejected back.

**Reasons.** Reject, decline and cancel ask why, and the other side sees it.

**Corrections, for Managers and Admins** (anyone who can act for any department):

- A received handover's quantity can be corrected, or the handover reversed.
- An element entry can be corrected or reversed.
- Nothing is deleted: the record keeps the original quantity, who changed it, when and why, and the audit log has both sides.
- Where the handover posted to Stock Management, the difference is posted there as an Adjustment, booked to the same department. Stock Management's no-negative rule still applies.
- A correction that would leave a department with fewer elements than it has already sent on or used is refused.

**Morning digest.** Each morning at about 7:05, each department's people get one mail listing what's waiting for them. Dispatch's mail also lists the orders due in the next three days, or late. Admins get the whole picture. Nobody is mailed when nothing is waiting for them. To turn it off, set the Script Property `PT_DIGEST` to `off`. `ADMIN_previewDigest()` shows what would be sent, without sending it.

**Security and housekeeping:**

- Only a Production Tracker sign-in is accepted here. A Stock Management or ELE Tracker session for the same person isn't.
- Ready for sign-in tokens stored hashed (`sha256:…`) in the Sessions tab.
- An unanswered request can be looked up only by the person who sent it.
- Typed text is trimmed and length-limited.
- File uploads commit before releasing their lock.
- Every failed request is logged in a new **ErrorLog** tab.
- File counts read two columns instead of the whole Attachments tab.
- A nightly job keeps the Idempotency log to 45 days and the ErrorLog to 90.
- The page has a Content-Security-Policy.

No new permissions. Production Tracker's switches in Team & Access are the same six.

### Deploying 3.2

1. **Backend.** Paste this Code.gs (3.2.0) and deploy it as a new version of the existing deployment (Deploy → Manage deployments → Edit → New version), so the URL stays the same.
2. Run **`ADMIN_setupSheet()`** once. It adds the new columns and the ErrorLog tab. Writes add any missing columns themselves too, so this step isn't urgent.
3. Run **`ADMIN_installTriggers()`** once. Google asks to approve sending mail and running daily jobs.
4. **Frontend.** Push index.html and sw.js. Phones get an update bar.

The 3.1 app keeps working against the 3.2 backend until the new page reaches everyone. The difference: reject, decline and cancel now need a reason.

## How a handover works

Material only moves between **neighbouring** departments on the route. Winding sends to Spray, or returns to Stores; it can't skip ahead to Short Clearance.

1. **Asked** (optional). The receiving department asks its neighbour to send something.
2. **Sent**. The sending department sends it, either answering a request or on its own.
3. **Received**. The receiving department confirms it has arrived. This is the moment custody changes.

At any point before receipt, a handover can be closed without moving anything:

- the receiver can **reject** what was sent
- the sender can **decline** a request
- whoever started it can **cancel** it

Sending *backwards* along the route is a **return**, and works exactly the same way. Returns are marked as such everywhere they appear.

## What reaches Stock Management

Raw material and finished goods live in Stock Management. A handover posts there **only when it is received**, and nothing is posted before that. The posting is made by the person receiving, booked to their own department, with the other department named in the remarks. Each handover carries a fixed posting ID, so a lost reply or a second press can't post twice. A rejected handover never posts at all.

| Handover | Posted in Stock Management as |
|---|---|
| Purchase → Stores (raw material) | Receipt into Stores |
| Stores → Winding (raw material) | Issue to Winding |
| Winding → Stores (return) | Return into Stores |
| Stores → Purchase (return to supplier) | Issue to Purchase |
| Assembly → Stores (finished goods) | Production into Stores |
| Stores → Dispatch | Dispatch, referencing the order if one was given |
| Dispatch → Stores (return) | Return into Stores |
| Stores → Assembly (rework) | Adjustment, taking it out of stock |

Elements between Winding and Assembly are not in Stock Management, and post nothing.

**Once this is live, those movements go through Production Tracker, not also through Stock Management.** Entering one in both places would count it twice.

## Elements (work in progress)

Elements are counted **by spec**. The specs are the ones Design defines in the **ELE_BOM** tab, plus any Winding starts when it makes the first of a new one.

A department's elements on hand, per spec, are:

> received + made − sent on − scrapped − used

Elements already sent count as gone, even before the other side confirms them, so they can't be sent twice. If the other side rejects them, they come back.

What each department can record is read off the route:

- **Winding** makes elements, since it receives raw material and sends elements.
- **Assembly** uses them up, since it receives elements and sends finished goods.
- **Any department holding elements** can scrap them.

## Files on handovers (3.1)

Any handover can carry photos and PDFs — Purchase's invoice with the film, a photo of rejected elements, a drawing with a request. Never required, up to 10 per handover.

- **On the Send and Ask forms**, *Add photos or PDFs* offers the camera, photos or files on a phone. Photos are shrunk on the phone first (about 2000 px on the long side), PDFs go as they are (up to 10 MB). The handover is saved first; its files follow in a tray under the form, and if one fails on a weak signal, **Retry files** sends just that one.
- **On every handover row**, **📎** shows how many files it has; tap it to see them, or **📎+** to add the first. Photos open full screen; PDFs open in a new tab on a computer or iPhone and are saved under their real name on an Android phone.
- **Who adds:** either department on the handover, or a Manager or Admin — when sending, asking, receiving, or later. Not on a handover that was rejected, declined or cancelled.
- **Who removes:** until it's received, the department that added a file; after, only a Manager or Admin. A removed file is kept in Drive and on record.

**Onto Stock Management.** When a raw-material or finished-goods handover is received, its files are recorded on the Stock Management entry the receipt posts — the same files, not copies — so Purchase's invoice is on the Receipt in Stock Management's history, marked *from HO-…*. Files added to a handover after its receipt follow it there. Removing a file from a received handover leaves it on the Stock Management entry; taking it off there is a Stock Management correction. Element handovers post nothing to Stock Management, so their files stay on the handover.

**Where the files are:** Google Drive, in the same **CALGAS Stock Attachments** folder Stock Management uses, under **Production Tracker**, a folder per month, each named after its handover. Never shared by link. The **Attachments** tab here records which file belongs to which handover and where it went.

## The route is a sheet

The **Workflow** tab has one row per link, with three columns: **FromDept**, **ToDept** and **Kind**. Kind is one of:

- **RM** — raw material
- **WIP** — elements
- **FG** — finished goods

`ADMIN_setupSheet()` fills it with the standard route. After that, it is yours to edit, and setup never overwrites it. Department names must match the **Departments** tab in Stock Management exactly. Raw material and finished goods are held by **Stores**.

## Who can do what

Access is managed in Stock Management, under **Team & Access**, on the **Production Tracker** switch.

| Feature | What it allows |
|---|---|
| Send and return from own department | Send, answer requests, cancel what you sent |
| Receive into own department | Receive, reject, ask a neighbour, cancel your requests |
| Act for any department | All of the above, for any department, and correcting or reversing entries (a Manager, by default) |
| Record produced, scrapped and used | Element entries in your own department |
| See every department | Look at any department's page (Management, by default) |
| Place orders | Placing, editing, closing and cancelling orders |

Admin holds everything. Marketing isn't a department on the route, so whoever takes orders needs **Place orders** granted to them personally.

## Moving over from requests

Run `ADMIN_migrateToHandovers()` once, after setup. Open raw-material requests become handovers:

- A request still waiting becomes a request to **Stores**.
- A request already **issued** was posted to Stock Management when it was issued. It comes across as Sent and marked as posted, so receiving it confirms custody without posting a second time.

Requests for Metallization or Slitting material are listed in the log, not moved. The function is safe to run again: each moved request is marked, and skipped next time.

## Deploying

1. **Stock Management first.** Deploy its Code.gs (3.11.0) as a new version, then run `ADMIN_migrateAccess()` once.
2. **Production Tracker backend.** Deploy this Code.gs (3.0.0) as a new version. Then run `ADMIN_setupSheet()`, which creates Workflow, Handovers, WIP_Entries and AuditLog, and fills the route. Then run `ADMIN_migrateToHandovers()`.
3. **Production Tracker frontend.** Push index.html and sw.js. Phones get an update bar.
4. Everyone signs in to Production Tracker again, so their session carries the new grants.

**Upgrading from 3.0 to 3.1 (files):** deploy Stock Management 3.14.0 first, then this Code.gs as a new version, and run **`ADMIN_setupAttachments()`** once from the editor. This app has never used Google Drive before, so Google will ask you to approve Drive access — approve it. The run makes the Attachments tab and folders and logs where files will go. Then push index.html and sw.js. If a receipt's files ever didn't reach Stock Management, running it again finishes them.

## Every change is logged

The **AuditLog** tab records who did what, and what it was before: every request, send, receipt, rejection, decline, cancellation, element entry and order.
