<img width="1672" height="941" alt="WSadder" src="https://github.com/user-attachments/assets/2ec5d5ec-4b5c-4a97-a56c-973c9f6d2c39" />

# WSadder - WhatsApp Group Adder

[![Node.js](https://img.shields.io/badge/Node.js-18%2B-339933?style=for-the-badge&logo=node.js&logoColor=white)](https://nodejs.org/)
[![npm](https://img.shields.io/badge/npm-install-CB3837?style=for-the-badge&logo=npm&logoColor=white)](https://www.npmjs.com/)
[![Excel](https://img.shields.io/badge/Excel-.xlsx-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.npmjs.com/package/xlsx)
[![WhatsApp Web](https://img.shields.io/badge/WhatsApp%20Web-unofficial-25D366?style=for-the-badge&logo=whatsapp&logoColor=white)](https://web.whatsapp.com/)
[![Responsible Use](https://img.shields.io/badge/Use-responsibly-blue?style=for-the-badge)](#responsible-use)

Project by Lay Yang (DKLY) - [www.dkly.net](https://www.dkly.net)

Repository: [github.com/osscv/WSadder](https://github.com/osscv/WSadder)

WSadder is a WhatsApp automation tool designed to simplify participant onboarding for event organizers, teams, communities, and individuals. It enables bulk adding of contacts from Excel directly into WhatsApp groups, reducing the need to manually add participants one by one.

Users can prepare a spreadsheet containing participant names and phone numbers, select a target WhatsApp group through WhatsApp Web, and let WSadder process the list efficiently. This helps reduce time, minimize human error, and keep group onboarding consistent when managing larger participant lists.

The tool also includes a smart fallback mechanism: if a participant cannot be added directly because of privacy settings, permissions, or restrictions, WSadder can automatically send a personalized direct message with the group invitation link. Meanwhile, the Excel file is updated with detailed status results such as added, invited, skipped, or failed, making tracking and follow-up easier.

WSadder is especially useful for workshops, training sessions, classes, volunteer programs, community events, and other organized activities where many confirmed participants need to be added to the correct WhatsApp group quickly and reliably.

This repository includes `sample-participants.xlsx` with fake participant data. Replace it with your own Excel file before running against real contacts.

## Features

- Bulk adding from Excel to WhatsApp groups
- Automated fallback invites via direct message
- Customizable invitation messages in `config.json`
- International phone number normalization
- Dry-run testing before real execution
- Status tracking with Excel updates
- Detailed run summary in `add_log.json`
- Built-in safeguards for responsible and consent-based use

## How It Works

1. Prepare an Excel file with participant names and WhatsApp phone numbers.
2. Run the setup wizard and select the Excel columns.
3. Scan the WhatsApp Web QR code.
4. Select the target WhatsApp group.
5. WSadder processes each participant.
6. If direct adding fails, WSadder can DM the participant with the group invite link.
7. The Excel file is updated with each participant's result.

## Use Cases

WSadder is designed for legitimate group onboarding workflows where participants have already registered, confirmed, or agreed to join a WhatsApp group. It works as a bulk automation tool that helps reduce repetitive manual work when moving large Excel participant lists into WhatsApp groups.

Common use cases include:

- Event organizers adding confirmed attendees to an event update group
- Workshop teams onboarding participants into session or announcement groups
- University clubs, student societies, and class coordinators managing class or activity groups
- Volunteer teams adding approved volunteers to operation or briefing groups
- Community managers moving registered members into the correct interest, project, or support group
- Training providers adding learners to course communication groups
- Nonprofit or campaign teams coordinating registered participants, helpers, or joiners

Instead of opening WhatsApp and adding each person manually, organizers can prepare an Excel file, select the target group, and let the tool process the list while writing the result back into the spreadsheet. When direct adding fails, the tool can use the configured invite message to DM the participant with the WhatsApp group link.

This is best used for small to medium operational workflows where the contact list is trusted, consent-based, and relevant to the group being managed. It is not intended for cold outreach, mass marketing, scraping, spam, or adding people who did not ask to be contacted.

## Responsible Use

This project is not affiliated with, endorsed by, sponsored by, or officially connected to WhatsApp, Meta, or any of their subsidiaries. It is an unofficial tool/plugin that uses WhatsApp Web automation.

Use this project only for legitimate communication with people who have agreed to be contacted or added to a group. Do not use this script for spam, harassment, scraping, scams, unauthorized marketing, impersonation, or any illegal activity.

You are responsible for complying with WhatsApp's terms, local laws, privacy rules, and consent requirements. Always verify that your participant list is accurate and that recipients expect to receive the group invite or message.

Using automation with WhatsApp may carry risks, including temporary limits, failed actions, account restrictions, or account suspension, especially if you abuse the tool, contact people without consent, send spam, or violate WhatsApp's terms or applicable laws. Use this project at your own risk.

The author is not responsible for bans, account restrictions, data loss, legal issues, misuse, or any damages caused by using this project.

## Requirements

- Node.js 18 or newer
- npm
- A WhatsApp account that can access the target group
- Admin permission in the target WhatsApp group is recommended, especially for adding participants and fetching invite links
- An Excel `.xlsx` file with a header row containing participant names and WhatsApp phone numbers

Check your installed Node.js and npm versions:

```bash
node -v
npm -v
```

Install Node.js:

| Platform | Recommended install method |
| --- | --- |
| Windows | Download the LTS installer from [nodejs.org](https://nodejs.org/) and run it |
| Ubuntu/Debian | Use NodeSource or your package manager, then install Node.js 18+ |
| macOS | Download from [nodejs.org](https://nodejs.org/) or install with Homebrew using `brew install node` |
| Other Linux | Use your distribution package manager, NodeSource, `nvm`, or another trusted Node.js installer |

Ubuntu/Debian example using NodeSource:

```bash
curl -fsSL https://deb.nodesource.com/setup_lts.x | sudo -E bash -
sudo apt-get install -y nodejs
```

## Install

Install the project dependencies from `package.json`:

```bash
npm install
```

This installs:

- `whatsapp-web.js` - WhatsApp Web automation client
- `qrcode-terminal` - prints the WhatsApp login QR code in the terminal
- `xlsx` - reads and updates Excel spreadsheets

If you need to install the dependencies manually, run:

```bash
npm install whatsapp-web.js qrcode-terminal xlsx
```

## Setup

1. Put your Excel `.xlsx` file in the project folder, or use the included `sample-participants.xlsx` for testing.
2. Run the setup wizard:

```bash
npm run setup
```

3. Choose the Excel file.
4. Choose the phone number column.
5. Choose the name column.
6. Enter the default country code, for example `60` for Malaysia.
7. Enter the event name and event date.
8. Scan the WhatsApp QR code when prompted.
9. Choose the target WhatsApp group.

The setup saves your choices in `config.json`.

By default, `config.json` has `"dryRun": true` so the sample project does not add or message anyone accidentally. Set it to `false` only when your real Excel file, WhatsApp group, and message are ready.

## Safety Checklist

Before running with real contacts:

- Confirm that participants agreed to be contacted or added to the group
- Run setup and verify the selected WhatsApp group
- Review `config.json`, especially `inviteMessage`, `eventName`, and `eventDate`
- Keep `"dryRun": true` for testing
- Set `"dryRun": false` only when you are ready for real adding or messaging
- Start with a small test list before processing a large Excel file

## Usage

Start the full add/invite run:

```bash
npm start
```

List WhatsApp groups and reselect the configured group:

```bash
npm run list-groups
```

Run setup again:

```bash
npm run setup
```

## Configuration

The main settings are stored in `config.json`:

| Field | Description |
| --- | --- |
| `excelFile` | Excel file to read and update |
| `phoneColumn` | Header name for the WhatsApp phone number column |
| `nameColumn` | Header name for the participant name column |
| `defaultCountryCode` | Country code added to local numbers, for example `60`, `65`, `SG`, or `Singapore` |
| `groupName` | Display name of the selected WhatsApp group |
| `groupId` | WhatsApp group ID saved after group selection |
| `eventName` | Event name used in the invite message |
| `eventDate` | Event date used in the invite message |
| `inviteMessage` | Customizable DM message template used when direct adding fails |
| `delayMsBetweenAdds` | Delay between each contact, in milliseconds |
| `dryRun` | Set to `true` to test without adding or messaging users |

### Custom DM Invite Message

When WSadder cannot add a participant directly to the group, it can send a fallback DM using the `inviteMessage` value in `config.json`.

You can customize the message text:

```json
{
  "inviteMessage": "Hi {name},\n\nYou're confirmed for {eventName} on {eventDate}. Please join the WhatsApp group here:\n{link}\n\nThank you!"
}
```

Available placeholders:

- `{name}` - participant name from Excel
- `{eventName}` - event name from `config.json`
- `{eventDate}` - event date from `config.json`
- `{link}` - WhatsApp group invitation link

## Phone Number Formats

Phone numbers are normalized before sending to WhatsApp. Supported formats include:

- Local numbers using `defaultCountryCode`, for example `01162383838`
- International numbers with `+`, for example `+60-1162383838`
- International numbers with separators, for example `+1 (415) 555-0101`
- International dialing prefix `00`, for example `00601162383838`

International numbers work for any country when the phone number includes the country code. Examples:

| Country/region | Example input | Normalized WhatsApp number |
| --- | --- | --- |
| Malaysia | `+60-1162383838` | `601162383838` |
| Singapore | `+65 9123 4567` | `6591234567` |
| China | `+86 138 0013 8000` | `8613800138000` |
| USA | `+1 (415) 555-0101` | `14155550101` |
| UK | `+44 7700 900123` | `447700900123` |
| Thailand | `+66 81 234 5678` | `66812345678` |
| Hong Kong | `+852 5123 4567` | `85251234567` |
| Taiwan | `+886 912 345 678` | `886912345678` |
| Vietnam | `+84 912 345 678` | `84912345678` |

For local numbers without `+` or `00`, set `defaultCountryCode` to the correct country code first. Supported country aliases include `MY`, `SG`, `CN`, `US`, `USA`, `UK`, `GB`, `TH`, `HK`, `TW`, `VN`, and full country names such as `Singapore`, `China`, `Thailand`, `Hong Kong`, `Taiwan`, and `Vietnam`.

Masked or incomplete values such as `+6011-2373-XXX` are skipped because WhatsApp requires a complete numeric phone number.

## Excel Output

During a run, the tool adds or updates a column named:

```text
WhatsApp Add Status
```

Each row is updated with the result for that participant, such as added, invited by DM, already in group, skipped, or failed.

The script also writes a run summary to:

```text
add_log.json
```

## Troubleshooting

| Issue | What to check |
| --- | --- |
| QR code does not work | Make sure WhatsApp Web is reachable and scan with the correct WhatsApp account |
| Group is not listed | Wait for WhatsApp Web sync, then run `npm run list-groups` again |
| Invite link cannot be fetched | Make sure your WhatsApp account has permission to access or generate the group invite link |
| Participants are skipped | Check the selected phone column and phone number format |
| Messages are not sent | Confirm `dryRun` is `false` and the recipient is registered on WhatsApp |
| Too many failures or limits | Stop the run and review WhatsApp account limits, consent, and message frequency |

## Git Notes

Do not commit local runtime files such as:

- `node_modules/`
- `.wwebjs_auth/`
- `.wwebjs_cache/`
- `add_log.json`

These are already listed in `.gitignore`.

## Notes

- Keep WhatsApp Web logged in while the script is running.
- Large WhatsApp accounts may take longer to sync group data.
- Use the tool responsibly and avoid adding or messaging people who did not consent to be contacted.
