# CALGAS Capacitors — Production Tracker 3.0

Material moving through the factory, one department to the next, and back when it has to go back.

```
Purchase → Stores → Winding → Spray → Short Clearance → Assembly-MFD / Assembly-kVAr → Stores → Dispatch
```

Metallization and Slitting are gone. Nothing is written for them any more, and their old records stay in the sheet as history.

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
| Act for any department | All of the above, for any department (a Manager, by default) |
| Record produced, scrapped and used | Element entries in your own department |
| See every department | Look at any department's page (Management, by default) |
| Place orders | The order form on the Overview |

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

## Every change is logged

The **AuditLog** tab records who did what, and what it was before: every request, send, receipt, rejection, decline, cancellation, element entry and order.
