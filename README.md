# Free E-Raffle

This project turns a donation export, volunteer list, ticket-pricing matrix, and Google Forms response export into an electronic raffle and a local draw-night presentation.

It is designed to be locally ran at the event site by someone who has no knowledge whatsoever of coding.

## The process is divided into four stages:

1. **Prepare** creates ticket balance and Forms unique code.
2. **Validate** sums Forms responses and compares to Prepare ticket balance file.
3. **Draw** creates winners from validation file (limitations of this project create need for separate winners file)
4. **Present** website that presents the winners from the secret results workbook.

`run_draw.bat` does not read or validate the raw Google Forms export. It reads only `raffle_validation.xlsx`.
The website does not read the validation workbook and does not calculate winners. It reads only `draw_data.js`, which `run_draw.bat` writes from the same data as `raffle_results.xlsx`.

---

## Project contents

Everything lives in one folder. There are no subfolders except the `.venv` created by `setup.bat`.

```text
\eRaffle
│
├── README.md
├── raffle.py
├── raffle_config.json
├── requirements.txt
│
├── setup.bat
├── run_prepare.bat
├── run_validate.bat
├── run_draw.bat
├── run_website.bat
├── draw.html
│
├── Clinic Fundraising_YYYY_MM_DD.csv
├── volunteer_list.xlsx
├── ticket_pricing.xlsx
├── googleform_results.xlsx (or googleform_results.csv)
├── manual_credits.xlsx
│
├── raffle_preparation.xlsx           created by run_prepare.bat
├── raffle_validation.xlsx            created by run_validate.bat
├── manual_basket_winners.xlsx        created by run_validate.bat, only if a basket has zero entries
├── raffle_results.xlsx               created by run_draw.bat; KEEP SECRET
└── draw_data.js                      created by run_draw.bat; KEEP SECRET
```

The included input files contain scrubbed demonstration data. You should run the complete example to understand the project's process and limitations.

---

# Setup Process

Utilize the setupfordummies.txt if unfamiliar with github repos, Python, or for the demonstration runthrough steps.

The setup only needs to be run once per project folder.
If the project folder is moved to another computer, run `setup.bat` again on that computer. New raffles must go into a new folder each time. When the raffle is over, you can delete the entire folder.

---

# Complete workflow for a new run

## Stage 0 — Start with a clean event folder

Before beginning a new raffle:

1. Requirements: Windows operating system, Python, chrome or edge browser
2. You must have a volunteer list, ticket pricing matrix and donations list in separate excel files that follow a specific naming convention and format. They must be in the project folder next to raffle.py. (Refer to the Input Examples branch for examples.)

### 1. Donations file

The default filename pattern is:

```text
Clinic Fundraising_*.csv
```

Only one file may match that pattern in the project folder. If two donation exports are present, the script stops instead of guessing which one is current. Any numbers, letters, etc. can be after the * in the file name.

The default expected columns are:

| Purpose | Default column |
|---|---|
| Donation amount | `Donation Amount` |
| Donation date | `Donation Date` |
| Donor first name | `Supporter First Name` |
| Donor last name | `Supporter Last Name` |
| Donor email | `Supporter Email` |
| Donation comment | `Donation Comment` |

These names can be changed in the `donation_columns` section of `raffle_config.json`.

### 2. Volunteer list

Default file:

```text
volunteer_list.xlsx
```

Default columns:

| Column | Required | Use |
|---|---:|---|
| `First` | Yes | Official first name |
| `Last` | Yes | Official last name |
| `Nickname` | No | Common nickname used by donors |
| `Aliases` | No | Extra aliases separated with semicolons or vertical bars |

Examples of aliases:

```text
DJ; D.J.; Doc J; DJacob
```

The script searches donation comments for:

- first and last name
- first name and last initial
- nickname and last name
- nickname and last initial
- aliases maintained in the workbook

A first name by itself is not used because it can credit the wrong volunteer. (Manually adjusted in later step.) There are three exceptions and one extra pass:

**Volunteers listed with only a first name.** If a volunteer's `Last` cell is empty, that exact first name (and nickname, if any) is the only possible way to match them, so it is used. `run_prepare.bat` prints a warning naming these volunteers because a bare first name in a comment has a higher chance of crediting the wrong person — give them a last name (even just an initial) if you can.

**Last names that are just an initial.** If the roster says `Michelle` / `S`, a comment saying `Michelle Shonka` still matches, because the roster initial is allowed to continue into a longer word. The reverse was already true (`Julia F` in a comment matches roster `Julia Fannin`). If two volunteers could both match, the donation is flagged for review instead of guessed.

**The donor's own name.** If the comment names no volunteer but the donor themself is on the volunteer list (a self-donation), the donation is credited automatically to the roster spelling of their name, with credit method `DONOR IS VOLUNTEER`. Volunteers do not need to write their own name in the comment when donating to themselves.

**Case merging.** Credits that differ only in capitalization or spacing ("michelle Shonka" vs "Michelle Shonka") are merged into one participant with one ticket code, preferring the roster spelling.

One more warning `run_prepare.bat` gives: if the volunteer list has extra columns containing text (for example, a scratch column of pasted names), it tells you those names are ignored. Anyone who should be matchable must be in the `First`/`Last` columns as their own row.

## 4. Ticket-pricing matrix

Default file:

```text
ticket_pricing.xlsx
```

Default sheet columns:

| Column | Meaning |
|---|---|
| `Donation` | Donation amount at a known pricing point |
| `Tickets` | Tickets awarded at that amount |

There can be unlimited donation amounts in the donation export. They do not need to appear exactly in the matrix.

The ticket estimate is calculated from the current pricing workbook every time `run_prepare.bat` runs:

- An exact matrix amount uses the exact ticket value.
- An amount between two rows uses piecewise-linear interpolation.
- An amount below the first row is estimated between `$0 = 0 tickets` and the first matrix row.
- An amount above the final row is extrapolated using the final two matrix rows.
- The result is rounded according to `ticket_estimation` in `raffle_config.json`.

The default setting rounds to the nearest whole ticket:

```json
"ticket_estimation": {
  "method": "piecewise_linear",
  "rounding_mode": "nearest",
  "round_to_increment": 1
}
```

Valid rounding modes are:
nearest (rounded to nearest whole number), floor (min), ceiling (max), round_to_increment, controls the ticket increment. (For example, 5 to rounds to multiples of five.)

Changes to `ticket_pricing.xlsx` automatically affect future preparation runs. No hard-coded ticket matrix is stored in the Python script.

---

# Stage 1 — Prepare

## After input files are formatted, double click `run_prepare.bat`.

### Input files used

`run_prepare.bat` reads (from the project folder):

- the donation CSV;
- `volunteer_list.xlsx`;
- `ticket_pricing.xlsx`;
- `manual_credits.xlsx`, (file created during `run_prepare.bat`)

It creates:

- `raffle_preparation.xlsx`;
- `manual_credits.xlsx` (updates this file after initial run)


### Edit Manual Credits

The first preparation run creates or refreshes:

```text
manual_credits.xlsx
```

The yellow columns are editable:

- `Manual Credit Name`
- `Manual Ticket Override`
- `Organizer Notes`

Donation credit follows this order:

1. A value entered in `Manual Credit Name` in `manual_credits.xlsx`.
2. One unambiguous volunteer found in the donation comment.
3. The donor’s displayed first and last name.
4. A generated unmatched-donation label if the donor name is also missing.

A donation is never discarded because the volunteer cannot be identified automatically. However, if there are donations that are unclaimed and will NOT be used in the raffle, you will need to put manually override those tickets to 0.

This means an unmatched donor still receives a ticket balance. If a volunteer later provides evidence that the donor intended to credit them, enter the volunteer’s official name in `Manual Credit Name`. Use `Manual Ticket Override` when the official ticket amount should intentionally differ from the pricing calculation. (MUST enter a whole number of tickets.)

Do not change the `Donation ID` column. It connects the manual decision to the source donation.

After entering corrections:

1. Save and close `manual_credits.xlsx`.
2. Run `run_prepare.bat`.
3. You can rinse and repeat as many times as you like. Once ALL donations are credited/accounted for, you can move on.

### Raffle Preparation

Review `raffle_preparation.xlsx`.

Important sheets:

#### `Summary`

Shows event totals and the number of comment matches needing review.

#### `Ticket Balances`

This is the official participant list used by validation.

It includes:

- participant name;
- total credited donations;
- official ticket balance;
- automatically generated ticket code;
- donation count;
- copyable invitation text.

Ticket codes are generated automatically.

If `raffle_preparation.xlsx` already exists, existing names retain their previous ticket codes where possible. Do not delete the preparation workbook after codes have been sent to participants unless you intend to rebuild and redistribute the codes.

### `Donation Detail`

Shows every donation, the selected credit name, the ticket-estimation method, interpolation basis, and final ticket award.

### `Pricing Matrix`

Shows the matrix used for that preparation run.

---

# Stage 2 — Build and distribute the Google Form

Create one Google Form for participant ticket allocation. The privacy setting must be open/public/anyone can take, otherwise participants will have to sign into Google (gross). Copy and paste the question titles below directly into the Form questions.

## Form Questions

### Ticket code

THIS MUST BE A REQUIRED RESPONSE

Question title:

```text
Your ticket code
```

The ticket code is authoritative. Participants should copy the code EXACTLY from the invitation text, including the hyphen. (Capitalization and extra spaces are normalized during validation.) The ticket code is what matches the ticket amount to the raffle participant, SO THEY MUST MATCH WHAT IS ON `raffle_preparation.xlsx`.

### Participant name

Question title:

```text
Your name
```

Included due to human error, easier for manual adjustments if needed.

### One numeric question per basket

The simplest approach: every basket question begins with:

```text
Tickets for:
```

Examples:

```text
Tickets for: Sunglasses
Tickets for: Chicago Coffee Basket
Tickets for: CORE VOLUNTEERS ONLY
```

The text after `Tickets for:` becomes the basket name on the website.

**If no question begins with that prefix**, the script falls back automatically: every column in the form export that is not the timestamp, name, or ticket-code column is treated as a basket, using the question text itself as the basket name. This means you can title your questions however you like (just the basket name, no prefix required) and it still works — you do not need to change `raffle_config.json` for this.

The one thing to watch for in fallback mode: Google Forms sometimes adds its own columns, most commonly `Email Address` when "Collect email addresses" is turned on. That column would otherwise look like a basket question with no valid ticket numbers in it, which would block the draw. It's excluded by default. If your form adds a different extra column (for example, a quiz "Score" column), add its exact header text to `ignore_columns` under `form` in `raffle_config.json`:

```json
"ignore_columns": ["Email Address", "Score"]
```

`run_validate.bat` prints the list of columns it decided to treat as baskets every time fallback mode is used, so you can check it caught the right ones before continuing.

Google Form settings:

- Use short-answer questions for each basket.
- Require a number greater than or equal to zero.
- A blank answer or zero means no tickets in that basket.
- Allow respondents to edit their responses and re-fill out the Form (later validation step only reads the latest submission.)

## Send the participant information

From `raffle_preparation.xlsx`, use the `Invite Text` column or send each participant:

- official ticket total
- ticket code
- Google Form link
- allocation deadline

Do not send the complete ticket-balance workbook to all participants because it contains everyone’s information. (I find it easiest to select the columns for name and invite text, print selection to PDF, and then hand them out and explain the process during handoff.)

---

# Stage 3 — Export and validate the Google Form

## Close the form

At the deadline:

1. Stop accepting responses.
2. Export the response sheet to Excel or CSV (Google Forms only exports to Google Sheets directly, but Google Sheets can export to either `.xlsx` or `.csv` from File > Download).
3. Save it inside the project folder as:

```text
googleform_results.xlsx
```
or
```text
googleform_results.csv
```

Keep only one of the two in the project folder at a time. If both are present, `run_validate.bat` stops and asks you to remove the older one, the same way it does for duplicate donation exports.

The workbook or CSV may have any worksheet/column order. For `.xlsx`, the first worksheet is read. Open and manually adjust any mis-typed ticket codes. Save and close the file before running validation.

## Run validation

Double-click:

```text
run_validate.bat
```

The script reads:

- `raffle_preparation.xlsx` for official names, ticket codes, and balances;
- `googleform_results.xlsx` or `googleform_results.csv` for basket allocations.

It creates:

```text
raffle_validation.xlsx
```
If needing to re-run `run_validate.bat` after later manual adjustments, you must delete the `raffle_validation.xlsx` file bc the script refuses to overwrite an existing validation workbook. This is intentional. It prevents an organizer from reviewing one version while the program silently replaces it with another. (No CHEATING!)

## Open `raffle_validation.xlsx`

### Validation sheets

#### `Validation Summary`

Shows counts by status and one overall draw status:

- `READY`
- `NOT READY - CORRECT THE FORM EXPORT AND REVALIDATE`

#### `Response Validation`

Contains one active row per ticket code and all basket allocations used by the next stage.

This is the sheet `run_draw.bat` reads.

#### `Basket Totals`

Shows each detected basket, its source column, submitted ticket total, draw-eligible total, and eligible entrant count.

#### `All Form Responses`

Includes all raw response rows, including older duplicates.

#### `Missing Responses`

Lists official ticket codes with no active form submission. A person with no response simply has no raffle entries; this does not block the draw.

### Allocation statuses

#### `EXACT`

Allocated tickets equal the official balance. The response is ready for the draw.

#### `UNDER`

Allocated tickets are lower than the official balance.

This response remains eligible by default. Any unused tickets do not count towards the draw.

#### `OVER`

Allocated tickets exceed the official balance.

The response is not silently excluded and is not automatically reduced. It remains visible in the validation workbook, and the entire draw is stopped until the source response is corrected and revalidated. This prevents the organizer from accidentally granting extra tickets or silently removing a participant.

Clarify with participant on whether mathematical error or intentional mistake. Dishonesty can be handled by voiding all tickets from form export.

#### `INVALID TICKET VALUE`

A basket contains a negative, fractional, or nonnumeric value. The draw is stopped until the source response is corrected.

#### `UNKNOWN CODE`

The code does not exist in `raffle_preparation.xlsx`. The draw is stopped because the response cannot be connected safely to an official balance.

#### `MISSING CODE`

No ticket code was entered. The draw is stopped.

#### `DUPLICATE - OLDER RESPONSE`

When one ticket code has multiple submissions, the latest response is used by default. Older submissions remain visible on `All Form Responses` but are not active. (For auditing/tracking purposes.)


## Correction and revalidation loop

When the validation workbook shows a problem:

1. Identify the participant and problem in `raffle_validation.xlsx`.
2. Correct the response in Google Forms or in the controlled response sheet, according to the nonprofit’s process.
3. Re-export the current responses to `googleform_results.xlsx` or `googleform_results.csv`.
4. Close Excel.
5. Delete `raffle_validation.xlsx`.
6. Run `run_validate.bat` again.
7. Repeat until the summary says the draw is ready.

Do not manually change `raffle_validation.xlsx` to force a draw. Correct the form source and create a fresh validation workbook. THE SITE WILL NOT WORK UNLESS THE SUMMARY SAYS THE DRAW IS READY.

---

# Stage 4 — Draw Winners

## Before drawing

Confirm that:

- the form is closed;
- `raffle_validation.xlsx` is the final reviewed version;
- the validation summary says `READY`;
- Excel is closed;
- `raffle_results.xlsx` does not already exist.

## Configure `raffle_config.jsn` if you want a less-generic name.

Update the name, date, and form link in `event` section:

```json
"event": {
  "name": "Volunteer Fundraising Raffle",
  "date": "2026",
  "form_link": ""
}
```

- `name` appears in generated workbooks and on the website.
- `date` appears in generated workbooks and on the website.
- `form_link` is appended to the invitation text when supplied.

#### Multiple winners policy

The default allows one person to win more than one basket. Edit `raffle_config.jsn` to prevent repeat winners by changing `true` to `false` before running `run_draw.bat`.

```json
"draw": {
  "allow_multiple_wins": true
}
```

When repeat winners are disabled, a prior winner’s later-basket entries are ignored. Make this rule clear to participants before collecting allocations.

#### Additional changes

If you would like to make any additional changes, that's up to you. I tried to make this as bland as humanly possible so it could have a wide variety of use cases. You'll have to edit the .html though.

###### Congratulations, you just made a website!
(Kinda)

## Run the draw

Double-click:

```text
run_draw.bat
```

This stage reads:

```text
raffle_validation.xlsx
raffle_preparation.xlsx
manual_basket_winners.xlsx    (only if it exists)
```

It does not read the donation export, pricing workbook, or Google Forms export directly. It checks every active validation row. If any row has `Ready for Draw = NO`, the draw stops. The person is not silently removed. Winner selection uses the positive ticket counts in the validation workbook. A participant with 50 tickets in a basket has five times the chance of a participant with 10 tickets in that basket. All zero and blank basket allocations are ignored. They are not written to the `Entries` sheet of `raffle_results.xlsx`.

A basket with zero positive tickets is skipped by default — unless you assign it manually. See **Assigning a winner to an empty basket** below.

The script creates:

```text
raffle_results.xlsx
```

This workbook is **secret until the live raffle**. **NO CHEATING**

### Assigning a winner to an empty basket

Sometimes nobody submits tickets for a basket — maybe it was added late, or nobody was interested at the price point, but the organizer still wants to give it to someone (a walk-in, a volunteer, a manual pick from a physical hat). There is no coding involved:

1. Run `run_validate.bat` as usual. If any basket has zero draw-eligible tickets, the script creates (or updates) `manual_basket_winners.xlsx` and lists it in the console output, one row per empty basket.
2. Open `manual_basket_winners.xlsx` in Excel.
3. In the yellow **Manual Winner (optional)** column, type the person's name **exactly** as it appears in `raffle_preparation.xlsx`'s `Participant Name` column. Capitalization doesn't need to match exactly, but the spelling does.
4. Save and close the file.
5. Leave the row blank for any basket you'd rather skip entirely — it will not appear in the results or on the website, exactly like today.
6. Run `run_draw.bat`.

If a name doesn't match anyone in `raffle_preparation.xlsx`, the draw stops before creating any secret files and tells you which basket and which typed name didn't match, so you can fix the spelling and try again.

On the website, a manually assigned basket skips the spinning-name animation (there's nothing to spin through) and reveals immediately, with a note that no tickets were submitted and the organizer assigned it directly. The `Results` sheet in `raffle_results.xlsx` marks these rows with `Manual Assignment = YES` so there's a permanent record of which winners came from the weighted draw versus a manual call.

If `manual_basket_winners.xlsx` doesn't exist yet, it means every basket currently has at least one draw-eligible ticket — nothing to do.

### Results workbook

`Event`
Contains event information and a confidentiality warning.

`Results`
Contains one winner per basket that has either draw-eligible tickets or a manual assignment. Includes a `Manual Assignment` column (`YES`/`NO`).

`Entries`
Contains only the positive ticket allocations used in the draw. Zero values are omitted. Manually assigned baskets have no rows here, since no tickets were submitted.

`run_draw.bat` also writes `draw_data.js`, which holds the same winners and entries in the format the website reads. The website uses the entries only for the spinning-name animation and displays the winner already computed by the draw. Like `raffle_results.xlsx`, `draw_data.js` is secret until the live raffle.

#### Preventing accidental redraws

`run_draw.bat` refuses to overwrite an existing `raffle_results.xlsx`.

This protects the first completed draw from being replaced accidentally.Delete `raffle_results.xlsx` only when the authorized organizer has deliberately decided that a completely new draw is required.

---

# Stage 5 — Present

## Time for the final .bat file!

```text
run_website.bat
```
This script:
1. checks that `draw.html` and `draw_data.js` exist;
2. starts a private web server bound to `127.0.0.1` on the current computer only;
3. opens the presentation in the default browser.

Keep the command window open during the presentation. Closing it stops the local website.

If the server cannot start (for example, port `8765` is already in use), double-clicking `draw.html` directly also works, because the page reads `draw_data.js` rather than the Excel workbook.

### Presentation controls

- Click **Draw winner** to animate and display the precomputed winner.
- Click **Next basket** to advance.
- Press `Space` or `Enter` to reveal or advance.

The browser loads all secret winners from `draw_data.js`, so the presentation computer and browser should be controlled by the event organizer.

### When the event is finished:

1. Close the browser tab.
2. Return to the command window.
3. Press `Ctrl+C` to stop the local server.
4. If applicable, store `raffle_results.xlsx` according to the nonprofit’s record-retention policy.
5. Pat yourself on the back for a job well done!

---

---

# Troubleshooting

## “The project is not set up yet”

Run `setup.bat` first.

## “More than one file matched”

More than one donation CSV matches the configured filename pattern. Move old exports out of the project folder so only the current file remains.

## “File is locked” or permission denied

Close the workbook in Excel and rerun the batch file.

## `raffle_validation.xlsx already exists`

Review whether it is the version you intend to keep. To validate a new form export, delete the old validation workbook and rerun `run_validate.bat`.

## `raffle_results.xlsx already exists`

The project is protecting secret winners from accidental replacement. Do not delete it unless an authorized new draw is required.

## No basket columns were found

This only happens when every column in the form export matched the timestamp, name, or ticket-code column, or was listed in `ignore_columns`, leaving nothing to treat as a basket. Open `googleform_results.xlsx`/`.csv` and confirm there is at least one other question column, and that `ignore_columns` in `raffle_config.json` isn't excluding a real basket by mistake.

## Manual winner does not match an official participant name

You typed a name in `manual_basket_winners.xlsx` that doesn't match anyone in the `Participant Name` column of `raffle_preparation.xlsx`. The draw stops before writing any secret files. Open `raffle_preparation.xlsx`, copy the person's name exactly as spelled there into the `Manual Winner (optional)` column, save, and rerun `run_draw.bat`. Capitalization can differ; spelling and spacing cannot.

## A typed name does not match

The ticket code is authoritative. A name difference is highlighted for review but does not block an otherwise valid response.

## A participant submitted twice

The latest response for that ticket code is used. Older responses remain visible on `All Form Responses`.

## The website says it could not load draw_data.js

- Confirm `draw_data.js` exists in the project folder. If not, run `run_draw.bat`.
- Confirm `draw.html` and `draw_data.js` are in the same folder.
- If using `run_website.bat`, keep the server command window open and confirm port `8765` is not already being used by another program. A previous server window that was never stopped with `Ctrl+C` is the usual cause.

## The browser shows old results

Close the browser tab and stop the server with `Ctrl+C`. Confirm the intended `raffle_results.xlsx` and `draw_data.js` are in the project folder, then run `run_website.bat` again.

## Ticket estimates look wrong

Review:

- `ticket_pricing.xlsx`;
- `ticket_estimation` in `raffle_config.json`;
- `Donation Detail` in `raffle_preparation.xlsx`.

The detail sheet shows the exact matrix points used for interpolation or extrapolation.

---

# Additional Notes

## Data Privacy

The files may contain donor names, emails, comments, volunteer names, ticket balances, and secret winners. The site is locally-hosted, so it is not exposed to other computers on the network. The .bat files open command prompts, and that can be scary! Make sure you trust that the author of this repo (Me) is not trying to steal sensitive personal information. (lowkey wish I was that talented/smart tbh)

## Intention

This repo is intended for public and free use. I tried to set it up in ways that would be easy to follow for anyone, and that it was mostly a "click and go" format. The requirements are minimal on purpose. The site style is as bland/boring as could possibly be so it could be applicable for a large use of cases.

### Thank you and good luck!! <3

---
