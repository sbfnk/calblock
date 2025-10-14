# calblock

Automatically blocks remaining free time in your Microsoft 365 calendar when it drops below a configurable threshold. Helps ensure you always have dedicated "silent work" time for focused, uninterrupted work.

## Features

- **Automatic blocking**: Monitors your calendar and creates "Silent work" blocks when free time drops below threshold
- **Configurable work hours**: Define your working day (e.g., 8:00-18:00)
- **Smart event filtering**: Only blocks based on accepted meetings and busy/tentative events
- **Existing silent work detection**: Accounts for already-blocked time
- **Focus reminders**: Optional daily reminders (e.g., "Check in on XZY" on Mondays)
- **Multiple-day processing**: Can process today or multiple days ahead

## Installation

This tool requires **M365-Auth** for OAuth2 authentication. If you don't have pipx, see the [M365-Auth installation guide](https://github.com/sbfnk/M365-Auth#installation).

```bash
# Install both packages
pipx install m365auth
pipx install calblock

# Set up authentication
get-token --profile calendar
```

**Note**: Both packages need to be installed separately because pipx creates isolated environments. Installing `calblock` includes the `m365auth` Python library, but you still need to install `m365auth` separately to get the `get-token` CLI command.

**Before PyPI publication**: Replace the install commands with:
```bash
git clone https://github.com/sbfnk/M365-Auth && pipx install ./M365-Auth
git clone https://github.com/sbfnk/calblock && pipx install ./calblock
```

## Configuration

On first run, calblock creates a default configuration file at `~/.config/calblock/config.yaml`:

```yaml
auth:
  profile: calendar  # M365 auth profile to use

calendar:
  api_base_url: https://graph.microsoft.com/beta
  timezone: Europe/London  # Automatically handles GMT/BST
  work_hours:
    start: 8
    end: 18
  min_free_hours: 2.0  # Minimum free + silent work time
  process_days:
    - 0  # Monday
    - 1  # Tuesday
    - 2  # Wednesday
    - 3  # Thursday
    - 4  # Friday
  silent_work_identifier: Silent work

focus_reminders:
  0: Check Project A    # Monday
  4: Check Project B    # Friday
```

Edit this file to customize:
- **work_hours**: Your typical working hours
- **min_free_hours**: Minimum combined free + silent work time (hours)
- **process_days**: Which days to process (0=Monday, 6=Sunday)
- **timezone**: Your timezone name
  - Examples: `Europe/London` (GMT/BST), `America/New_York` (EST/EDT), `Asia/Tokyo`, `UTC`
  - Full list: https://en.wikipedia.org/wiki/List_of_tz_database_time_zones
- **silent_work_identifier**: Title for blocking events
- **focus_reminders**: Optional day-specific reminders

## Usage

```bash
# Process today (default)
calblock

# Process next 5 days
calblock --days 5

# Process specific date
calblock --date 2025-10-15

# Verbose mode (shows detailed calculations)
calblock -v

# Custom config file
calblock --config ~/my-config.yaml
```

## How It Works

1. **Fetch events**: Retrieves your calendar events for the specified date
2. **Calculate free time**:
   - Identifies busy periods (accepted meetings, busy/tentative events)
   - Calculates free slots within work hours
   - Accounts for existing "Silent work" blocks
3. **Check threshold**: If `free time + existing silent work ≤ min_free_hours`
4. **Create blocks**: Fills free slots with "Silent work" events until threshold is met

### Event Filtering

The blocker intelligently filters events:
- **Included**: Accepted meetings, busy events, tentative events, out-of-office
- **Excluded**: Declined/not-responded invitations, "free" status, "working elsewhere"

This ensures you don't over-block based on tentative or unconfirmed time.

## Example Output

```
2025-10-13: 1.5h free + 0.5h silent = 2.0h total
✅ No blocking needed

2025-10-14: 0.5h free + 0.0h silent = 0.5h total
Created: Silent work (Check in on research project) 14:00-15:30 (90min)
✅ Blocked 1 slots

2025-10-15: 3.0h free + 0.0h silent = 3.0h total
✅ No blocking needed
```

## Scheduling

Run automatically via cron:

```bash
# Add to crontab (crontab -e)
# Run at 7am every weekday, check next 2 days
0 7 * * 1-5 /Users/yourusername/.local/bin/calblock --days 2
```

## Authentication

calblock uses the `calendar` profile from M365-Auth. Ensure you've run:

```bash
get-token --profile calendar
```

The OAuth refresh token is stored securely in your system keychain and automatically refreshed as needed.

## Troubleshooting

**"Error: m365auth module not found"**
- Install M365-Auth: https://github.com/sbfnk/M365-Auth#installation

**"No refresh token found"**
- Run: `get-token --profile calendar`

**"Events not being filtered correctly"**
- Use `--verbose` mode to see detailed event processing
- Check the `showAs` status and response status of events

**"Wrong timezone"**
- Set the correct timezone name in config.yaml (e.g., `timezone: America/New_York`)
- Use `--verbose` to see how times are being interpreted

## License

MIT License

## Credits

Built with [M365-Auth](https://github.com/sbfnk/M365-Auth) for OAuth2 authentication.
