import os
import json
import time
from datetime import datetime, timedelta

CACHE_FILE = os.path.join(os.path.expanduser("~"), ".edge_tts_gui_voices_cache.json")
CACHE_EXPIRY_DAYS = 7  # Cache expires after 7 days

def format_timestamp(timestamp):
    """Convert timestamp to human readable format"""
    return datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')

def get_cache_status():
    """Get detailed cache status information"""
    try:
        if not os.path.exists(CACHE_FILE):
            return {
                'exists': False,
                'message': 'No cache file found',
                'last_updated': None,
                'expires_in': None
            }
            
        with open(CACHE_FILE, 'r') as f:
            cache_data = json.load(f)
            
        timestamp = cache_data.get('timestamp', 0)
        current_time = time.time()
        expiry_time = timestamp + (CACHE_EXPIRY_DAYS * 24 * 60 * 60)
        time_until_expiry = expiry_time - current_time
        
        if time_until_expiry <= 0:
            return {
                'exists': True,
                'message': 'Cache expired',
                'last_updated': format_timestamp(timestamp),
                'expires_in': 'Expired',
                'voice_count': len(cache_data.get('voices', []))
            }
        
        days_left = int(time_until_expiry / (24 * 60 * 60))
        hours_left = int((time_until_expiry % (24 * 60 * 60)) / 3600)
        
        return {
            'exists': True,
            'message': 'Cache valid',
            'last_updated': format_timestamp(timestamp),
            'expires_in': f'{days_left}d {hours_left}h',
            'voice_count': len(cache_data.get('voices', []))
        }
    except Exception as e:
        return {
            'exists': False,
            'message': f'Error reading cache: {str(e)}',
            'last_updated': None,
            'expires_in': None
        }

def load_cached_voices():
    """Load voices from cache if available and not expired"""
    try:
        if not os.path.exists(CACHE_FILE):
            return None
            
        with open(CACHE_FILE, 'r') as f:
            cache_data = json.load(f)
            
        # Check if cache is expired
        cache_timestamp = cache_data.get('timestamp', 0)
        if time.time() - cache_timestamp > (CACHE_EXPIRY_DAYS * 24 * 60 * 60):
            return None
            
        return cache_data.get('voices')
    except Exception:
        return None

def save_voices_to_cache(voices):
    """Save voices list to cache with current timestamp"""
    try:
        cache_data = {
            'timestamp': time.time(),
            'voices': voices
        }
        
        os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
        with open(CACHE_FILE, 'w') as f:
            json.dump(cache_data, f)
        return True
    except Exception as e:
        return False 