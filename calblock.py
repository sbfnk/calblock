#!/usr/bin/env python3
"""
Configurable Calendar Auto-Blocker
Automatically blocks remaining free time when it drops below threshold
"""

import datetime
import requests
import json
import yaml
from typing import List, Dict, Tuple, Optional
import argparse
import os
import sys
from pathlib import Path
from dateutil.parser import parse as parse_date

# Timezone support (zoneinfo in Python 3.9+, fallback to pytz)
try:
    from zoneinfo import ZoneInfo
except ImportError:
    from backports.zoneinfo import ZoneInfo

# Import M365 auth library
try:
    from m365auth import get_access_token
except ImportError:
    print("Error: m365auth module not found. Please ensure m365auth is installed.")
    print("You can install it with: pip install -e /path/to/M365-Auth")
    sys.exit(1)

# XDG Base Directory support
def get_xdg_config_home() -> Path:
    """Get XDG config directory"""
    xdg_config_home = os.environ.get('XDG_CONFIG_HOME')
    if xdg_config_home:
        return Path(xdg_config_home)
    else:
        return Path.home() / '.config'

def get_default_config_path() -> Path:
    """Get default config file path following XDG spec"""
    return get_xdg_config_home() / 'calblock' / 'config.yaml'

class CalendarBlocker:
    def __init__(self, config_path: Optional[str] = None, verbose: bool = False):
        if config_path is None:
            config_path = get_default_config_path()
        else:
            config_path = Path(config_path).expanduser()
        
        self.config = self.load_config(config_path)
        self.access_token = None
        self.verbose = verbose
        
    def vprint(self, *args, **kwargs):
        """Print only if verbose mode is enabled"""
        if self.verbose:
            print(*args, **kwargs)
            
    def load_config(self, config_path: Path) -> dict:
        """Load configuration from YAML file"""
        
        # Default configuration
        default_config = {
            'auth': {
                'profile': 'calendar'  # M365 auth profile to use
            },
            'calendar': {
                'api_base_url': 'https://graph.microsoft.com/beta',
                'timezone': 'Europe/London',
                'work_hours': {
                    'start': 8,
                    'end': 18
                },
                'min_free_hours': 2.0,
                'process_days': [0, 1, 2, 3, 4],  # Monday=0 to Friday=4
                'silent_work_identifier': 'Silent work'
            },
            'focus_reminders': {
                0: 'Check CMMID',    # Monday
                4: 'Check HPRU'      # Friday
            }
        }
        
        if not config_path.exists():
            # Create default config file
            config_path.parent.mkdir(parents=True, exist_ok=True)
            with open(config_path, 'w') as f:
                yaml.dump(default_config, f, default_flow_style=False, sort_keys=False)
            print(f"Created default config at {config_path}")
            print("Please edit this file with your settings and run again.")
            sys.exit(1)
        
        try:
            with open(config_path, 'r') as f:
                config = yaml.safe_load(f)
            
            # Merge with defaults for missing keys
            def merge_dicts(default, user):
                for key, value in default.items():
                    if key not in user:
                        user[key] = value
                    elif isinstance(value, dict) and isinstance(user[key], dict):
                        merge_dicts(value, user[key])
            
            merge_dicts(default_config, config)
            return config
            
        except Exception as e:
            print(f"❌ Error loading config from {config_path}: {e}")
            sys.exit(1)
    
    def get_access_token(self) -> str:
        """Get OAuth token using m365auth library"""
        try:
            profile = self.config['auth'].get('profile', 'calendar')
            self.vprint(f"🔑 Getting access token for profile: {profile}")
            self.access_token = get_access_token(profile)
            self.vprint(f"🔑 Got access token (length: {len(self.access_token)})")
            return self.access_token
        except Exception as e:
            print(f"❌ Failed to get access token: {e}")
            raise
    
    def get_timezone(self) -> ZoneInfo:
        """Get timezone from config"""
        tz_name = self.config['calendar'].get('timezone', 'UTC')
        tz = ZoneInfo(tz_name)
        self.vprint(f"🌐 Using timezone: {tz_name}")
        return tz
    
    def should_process_date(self, date: datetime.date) -> bool:
        """Check if we should process this date based on config"""
        should_process = date.weekday() in self.config['calendar']['process_days']
        day_name = date.strftime('%A')
        self.vprint(f"📅 {date} ({day_name}): {'✓' if should_process else '✗'} process (weekday {date.weekday()})")
        return should_process
    
    def get_calendar_events(self, date: datetime.date) -> List[Dict]:
        """Get calendar events for a specific date"""
        if not self.access_token:
            self.get_access_token()
        
        tz = self.get_timezone()
        start_time = datetime.datetime.combine(date, datetime.time(0, 0, 0), tzinfo=tz)
        end_time = datetime.datetime.combine(date, datetime.time(23, 59, 59), tzinfo=tz)
        
        self.vprint(f"📊 Fetching events from {start_time.isoformat()} to {end_time.isoformat()}")
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        base_url = self.config['calendar']['api_base_url']
        url = f"{base_url}/me/calendar/calendarView"
        params = {
            'startDateTime': start_time.isoformat(),
            'endDateTime': end_time.isoformat(),
            '$select': 'subject,start,end,showAs,organizer,responseStatus',
            '$orderby': 'start/dateTime',
            '$top': 100
        }
        
        self.vprint(f"📊 API URL: {url}")
        self.vprint(f"📊 Query params: {params}")
        
        response = requests.get(url, params=params, headers=headers, timeout=30)
        response.raise_for_status()
        
        events = response.json().get('value', [])
        self.vprint(f"📊 Retrieved {len(events)} events from API")
        
        return events
    
    def is_silent_work(self, event_subject: str) -> bool:
        """Check if event is silent work based on config"""
        identifier = self.config['calendar']['silent_work_identifier'].lower()
        is_silent = identifier in event_subject.lower()
        self.vprint(f"🔍 '{event_subject}' contains '{identifier}': {is_silent}")
        return is_silent
    
    def calculate_free_time(self, events: List[Dict], date: datetime.date) -> Tuple[int, List[Tuple[datetime.datetime, datetime.datetime]], int]:
        """Calculate free time and return free slots + existing silent work minutes"""
        tz = self.get_timezone()
        work_start_hour = self.config['calendar']['work_hours']['start']
        work_end_hour = self.config['calendar']['work_hours']['end']
        
        # Convert decimal hours to time objects
        start_hour = int(work_start_hour)
        start_minute = int((work_start_hour - start_hour) * 60)
        end_hour = int(work_end_hour)
        end_minute = int((work_end_hour - end_hour) * 60)
        
        work_start = datetime.datetime.combine(date, datetime.time(start_hour, start_minute), tzinfo=tz)
        work_end = datetime.datetime.combine(date, datetime.time(end_hour, end_minute), tzinfo=tz)
        
        self.vprint(f"⏰ Work hours: {work_start.strftime('%H:%M')} - {work_end.strftime('%H:%M')}")

        # Define which showAs values should block time
        blocking_statuses = {'busy', 'oof', 'tentative'}  # oof = out of office
        non_blocking_statuses = {'free', 'workingelsewhere'}  # workingElsewhere shouldn't block focus time
    
        busy_periods = []
        existing_silent_work_minutes = 0
        skipped_events = []
        processed_events = []
        
        for i, event in enumerate(events):
            subject = event.get('subject', 'Untitled')
            show_as = event.get('showAs', '').lower()
            response_status = event.get('responseStatus', {}).get('response', 'none')
            
            self.vprint(f"\n📋 Event {i+1}: '{subject}'")
            self.vprint(f"   ShowAs: {show_as}")
            self.vprint(f"   Response: {response_status}")

            # Skip non-blocking showAs statuses
            if show_as in non_blocking_statuses:
                self.vprint(f"   ❌ SKIPPED: ShowAs '{show_as}' is non-blocking")
                skipped_events.append(f"'{subject}' - showAs: {show_as} (non-blocking)")
                continue

            # Skip events that aren't explicitly blocking (unless they're accepted meetings)
            if show_as not in blocking_statuses:
                # For unknown showAs values, fall back to checking if it's an accepted meeting
                if response_status.lower() not in ['accepted', 'organizer']:
                    self.vprint(f"   ❌ SKIPPED: ShowAs '{show_as}' not explicitly blocking and not accepted")
                    skipped_events.append(f"'{subject}' - showAs: {show_as}, response: {response_status}")
                    continue
                else:
                    self.vprint(f"   ⚠️  Unknown showAs '{show_as}' but treating as blocking (accepted meeting)")
                   
            # Skip non-accepted invitations for explicitly blocking events
            if show_as in blocking_statuses and response_status.lower() not in ['accepted', 'organizer']:
                self.vprint(f"   ❌ SKIPPED: Response status '{response_status}' not accepted/organizer")
                skipped_events.append(f"'{subject}' - response status: {response_status}")
                continue
            
            start = parse_date(event['start']['dateTime'])
            end = parse_date(event['end']['dateTime'])
            
            self.vprint(f"   Start: {start}")
            self.vprint(f"   End: {end}")
            
            if start.tzinfo is None:
                start = start.replace(tzinfo=datetime.timezone.utc)
                self.vprint(f"   ⚠️  Added UTC timezone to start time")
            if end.tzinfo is None:
                end = end.replace(tzinfo=datetime.timezone.utc)
                self.vprint(f"   ⚠️  Added UTC timezone to end time")
            
            start = start.astimezone(tz)
            end = end.astimezone(tz)
            
            self.vprint(f"   Local Start: {start}")
            self.vprint(f"   Local End: {end}")
            
            # Check if event is silent work
            if self.is_silent_work(subject):
                work_start_time = max(start, work_start)
                work_end_time = min(end, work_end)
                if work_start_time < work_end_time:
                    silent_minutes = (work_end_time - work_start_time).total_seconds() / 60
                    existing_silent_work_minutes += silent_minutes
                    self.vprint(f"   📝 Silent work: {silent_minutes:.1f} minutes in work hours")
                else:
                    self.vprint(f"   📝 Silent work: Outside work hours")
            
            # Clip to work hours for busy period calculation
            clipped_start = max(start, work_start)
            clipped_end = min(end, work_end)
            
            if clipped_start < clipped_end:
                duration_minutes = (clipped_end - clipped_start).total_seconds() / 60
                busy_periods.append((clipped_start, clipped_end))
                self.vprint(f"   ✅ PROCESSED: {clipped_start.strftime('%H:%M')}-{clipped_end.strftime('%H:%M')} ({duration_minutes:.1f}min)")
                processed_events.append(f"'{subject}' - {clipped_start.strftime('%H:%M')}-{clipped_end.strftime('%H:%M')}")
            else:
                self.vprint(f"   ❌ SKIPPED: Outside work hours")
                skipped_events.append(f"'{subject}' - outside work hours")
        
        # Print summary of event processing
        self.vprint(f"\n📊 Event Processing Summary:")
        self.vprint(f"   Total events from API: {len(events)}")
        self.vprint(f"   Processed events: {len(processed_events)}")
        self.vprint(f"   Skipped events: {len(skipped_events)}")
        
        if processed_events:
            self.vprint(f"\n✅ Processed Events:")
            for event in processed_events:
                self.vprint(f"   - {event}")
        
        if skipped_events:
            self.vprint(f"\n❌ Skipped Events:")
            for event in skipped_events:
                self.vprint(f"   - {event}")
        
        # Merge overlapping periods
        busy_periods.sort()
        self.vprint(f"\n🔄 Merging {len(busy_periods)} busy periods...")
        
        merged_periods = []
        for start, end in busy_periods:
            if merged_periods and start <= merged_periods[-1][1]:
                old_end = merged_periods[-1][1]
                merged_periods[-1] = (merged_periods[-1][0], max(merged_periods[-1][1], end))
                self.vprint(f"   🔗 Merged with previous: extended to {merged_periods[-1][1].strftime('%H:%M')}")
            else:
                merged_periods.append((start, end))
                self.vprint(f"   ➕ New period: {start.strftime('%H:%M')}-{end.strftime('%H:%M')}")
        
        # Calculate free slots
        free_slots = []
        current_time = work_start
        
        self.vprint(f"\n🆓 Calculating free slots...")
        
        for i, (busy_start, busy_end) in enumerate(merged_periods):
            if current_time < busy_start:
                slot_duration = (busy_start - current_time).total_seconds() / 60
                free_slots.append((current_time, busy_start))
                self.vprint(f"   ✅ Free slot {len(free_slots)}: {current_time.strftime('%H:%M')}-{busy_start.strftime('%H:%M')} ({slot_duration:.1f}min)")
            else:
                self.vprint(f"   ❌ No gap before busy period {i+1}")
            current_time = max(current_time, busy_end)
        
        # Final free slot
        if current_time < work_end:
            slot_duration = (work_end - current_time).total_seconds() / 60
            free_slots.append((current_time, work_end))
            self.vprint(f"   ✅ Final free slot: {current_time.strftime('%H:%M')}-{work_end.strftime('%H:%M')} ({slot_duration:.1f}min)")
        
        total_free_minutes = sum((end - start).total_seconds() / 60 for start, end in free_slots)
        
        self.vprint(f"\n📊 Time Calculation Results:")
        self.vprint(f"   Total free minutes: {total_free_minutes:.1f}")
        self.vprint(f"   Existing silent work minutes: {existing_silent_work_minutes:.1f}")
        self.vprint(f"   Number of free slots: {len(free_slots)}")
        
        return int(total_free_minutes), free_slots, int(existing_silent_work_minutes)
    
    def create_blocking_events(self, free_slots: List[Tuple[datetime.datetime, datetime.datetime]], 
                             date: datetime.date, minutes_needed: int) -> List[Dict]:
        """Create silent work blocking events for exactly the minutes needed"""
        if not self.access_token:
            self.get_access_token()
        
        if minutes_needed <= 0:
            return []
        
        self.vprint(f"\n🔨 Creating blocking events for {minutes_needed} minutes...")
        
        created_events = []
        minutes_blocked = 0
        
        # Get focus reminder for this day
        focus_note = self.config['focus_reminders'].get(date.weekday(), "")
        silent_work_title = self.config['calendar']['silent_work_identifier']
        
        for i, (start, end) in enumerate(free_slots):
            if minutes_blocked >= minutes_needed:
                break
                
            slot_minutes = (end - start).total_seconds() / 60
            self.vprint(f"   Slot {i+1}: {start.strftime('%H:%M')}-{end.strftime('%H:%M')} ({slot_minutes:.1f}min)")
            
            if slot_minutes < 15:
                self.vprint(f"   ❌ Too short (< 15min), skipping")
                continue
            
            remaining_needed = minutes_needed - minutes_blocked
            minutes_to_use = min(slot_minutes, remaining_needed)
            block_end = start + datetime.timedelta(minutes=minutes_to_use)
            
            self.vprint(f"   ✅ Using {minutes_to_use:.1f} minutes: {start.strftime('%H:%M')}-{block_end.strftime('%H:%M')}")
            
            event_data = {
                'subject': silent_work_title,
                'start': {
                    'dateTime': start.isoformat(),
                    'timeZone': start.tzinfo.tzname(None) or 'UTC'
                },
                'end': {
                    'dateTime': block_end.isoformat(),
                    'timeZone': block_end.tzinfo.tzname(None) or 'UTC'
                },
                'showAs': 'busy',
                'categories': [silent_work_title],
                'isReminderOn': False
            }
            
            if focus_note:
                event_data['body'] = {
                    'contentType': 'text',
                    'content': focus_note
                }
            
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Content-Type': 'application/json'
            }
            
            base_url = self.config['calendar']['api_base_url']
            self.vprint(f"   📡 Creating event via API...")
            
            response = requests.post(f"{base_url}/me/calendar/events", json=event_data, headers=headers)
            
            if response.status_code == 201:
                created_events.append(response.json())
                note_text = f" ({focus_note})" if focus_note else ""
                print(f"Created: {silent_work_title}{note_text} {start.strftime('%H:%M')}-{block_end.strftime('%H:%M')} ({minutes_to_use:.0f}min)")
                minutes_blocked += minutes_to_use
                self.vprint(f"   ✅ Event created successfully")
            else:
                self.vprint(f"   ❌ Failed to create event: {response.status_code} {response.text}")
        
        return created_events
    
    def process_date(self, date: datetime.date) -> bool:
        """Process a single date"""
        if not self.should_process_date(date):
            return False
        
        self.vprint(f"\n🔍 Processing {date}...")
        
        events = self.get_calendar_events(date)
        free_minutes, free_slots, existing_silent_minutes = self.calculate_free_time(events, date)
        
        free_hours = free_minutes / 60
        existing_silent_hours = existing_silent_minutes / 60
        total_available = free_hours + existing_silent_hours
        min_free_hours = self.config['calendar']['min_free_hours']
        
        print(f"{date.strftime('%Y-%m-%d')}: {free_hours:.1f}h free + {existing_silent_hours:.1f}h silent = {total_available:.1f}h total")
        
        self.vprint(f"   Minimum required: {min_free_hours:.1f}h")
        self.vprint(f"   Total available: {total_available:.1f}h")
        self.vprint(f"   Need blocking: {total_available <= min_free_hours}")
        
        if total_available <= min_free_hours:
            min_silent_minutes = min_free_hours * 60
            additional_minutes_needed = max(0, min_silent_minutes - existing_silent_minutes)
            
            self.vprint(f"   Additional minutes needed: {additional_minutes_needed:.1f}")
            
            if additional_minutes_needed > 0:
                created_events = self.create_blocking_events(free_slots, date, additional_minutes_needed)
                if created_events:
                    print(f"✅ Blocked {len(created_events)} slots")
                    return True
                else:
                    print(f"❌ Need {additional_minutes_needed}min more but no free slots available")
                    return False
            else:
                print(f"✅ Already sufficient silent work")
                return False
        else:
            print(f"✅ No blocking needed")
            return False
    
    def run_date_range(self, start_date: datetime.date, days: int = 1) -> None:
        """Run the blocker for a date range"""
        blocked_days = 0
        
        for i in range(days):
            check_date = start_date + datetime.timedelta(days=i)
            try:
                if self.process_date(check_date):
                    blocked_days += 1
            except Exception as e:
                print(f"❌ Error on {check_date}: {e}")
                if self.verbose:
                    import traceback
                    traceback.print_exc()
        
        if blocked_days > 0:
            print(f"📊 Blocked time on {blocked_days} day(s)")

def main():
    parser = argparse.ArgumentParser(description="Configurable Calendar Auto-Blocker")
    
    date_group = parser.add_mutually_exclusive_group()
    date_group.add_argument('--date', '-d', type=str, help='Specific date to process (YYYY-MM-DD)')
    date_group.add_argument('--days', '-n', type=int, default=1, help='Number of days to process from today (default: 1)')
    
    parser.add_argument('--config', '-c', type=str, 
                        help=f'Path to config file (default: {get_default_config_path()})')
    parser.add_argument('--verbose', '-v', action='store_true', 
                        help='Enable verbose debug output')
    
    args = parser.parse_args()
    
    # Determine date range
    if args.date:
        try:
            start_date = datetime.datetime.strptime(args.date, '%Y-%m-%d').date()
            days_to_process = 1
        except ValueError:
            print("❌ Invalid date format. Use YYYY-MM-DD")
            return
    else:
        start_date = datetime.date.today()
        days_to_process = args.days
    
    # Initialize and run blocker
    try:
        blocker = CalendarBlocker(config_path=args.config, verbose=args.verbose)
        blocker.run_date_range(start_date, days_to_process)
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()
