"""User-local keyword profiles. No profile content is embedded in the application."""
import json
import os
import re
import tempfile
import uuid
from pathlib import Path

import document_cleanup_core as core

SCHEMA_VERSION = 1
MAX_PROFILE_BYTES = 256 * 1024


class ProfileError(ValueError):
    pass


def user_directory():
    base = Path(os.environ['APPDATA']) if os.environ.get('APPDATA') else Path.home() / '.config'
    return base / 'DocumentScanCleanup'


def validate_profile(data):
    if not isinstance(data, dict) or type(data.get('schema_version')) is not int or data['schema_version'] != SCHEMA_VERSION:
        raise ProfileError('不支援的清單格式／版本。')
    name = data.get('name')
    if not isinstance(name, str) or not name.strip() or len(name.strip()) > 80 or any(ord(c) < 32 for c in name):
        raise ProfileError('清單名稱須為 1～80 字，不可含換行或控制字元。')
    words = data.get('keywords')
    if not isinstance(words, list) or not all(isinstance(word, str) for word in words):
        raise ProfileError('keywords 必須是文字陣列。')
    if any(any(c in word for c in ';\r\n\x00') for word in words):
        raise ProfileError('單一關鍵字不可含分號、換行或空字元。')
    try:
        _, words = core.pattern(words)
    except core.DocumentError as exc:
        raise ProfileError(str(exc)) from exc
    return {'schema_version': SCHEMA_VERSION, 'name': name.strip(), 'keywords': words}


def read_json(path, limit=MAX_PROFILE_BYTES):
    try:
        with Path(path).open('rb') as stream:
            raw = stream.read(limit + 1)
        if len(raw) > limit:
            raise ProfileError('設定檔過大。')
        return json.loads(raw.decode('utf-8-sig'))
    except (OSError, UnicodeError, ValueError) as exc:
        raise ProfileError('無法讀取有效的 UTF-8 JSON 設定檔。') from exc


def load_profile(path):
    return validate_profile(read_json(path))


def atomic_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', newline='\n', dir=path.parent,
                                         prefix='.tmp-', suffix='.json', delete=False) as stream:
            temporary = Path(stream.name)
            json.dump(data, stream, ensure_ascii=False, indent=2)
            stream.write('\n')
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


class ProfileStore:
    def __init__(self, directory=None):
        self.directory = Path(directory) if directory is not None else user_directory()
        self.profiles = self.directory / 'profiles'
        self.settings = self.directory / 'keyword-settings.json'

    def profile_path(self, identifier):
        if not isinstance(identifier, str) or not re.fullmatch(r'[a-f0-9]{32}', identifier):
            raise ProfileError('無效的本機清單識別碼。')
        return self.profiles / (identifier + '.keywords.json')

    def load(self, identifier):
        return load_profile(self.profile_path(identifier))

    def save(self, name, words, identifier=None):
        profile = validate_profile({'schema_version': SCHEMA_VERSION, 'name': name, 'keywords': words})
        identifier = identifier or uuid.uuid4().hex
        atomic_json(self.profile_path(identifier), profile)
        return identifier, profile

    def import_profile(self, path):
        profile = load_profile(path)
        return self.save(profile['name'], profile['keywords'])

    def list_profiles(self):
        entries, skipped = [], 0
        for path in sorted(self.profiles.glob('*.keywords.json')):
            identifier = path.name.removesuffix('.keywords.json')
            try:
                profile = self.load(identifier)
                entries.append((identifier, profile))
            except ProfileError:
                skipped += 1
        return entries, skipped

    def set_startup(self, identifier):
        if identifier is not None:
            self.load(identifier)
        atomic_json(self.settings, {'schema_version': SCHEMA_VERSION, 'startup_profile': identifier})

    def startup(self):
        generic = {'schema_version': SCHEMA_VERSION, 'name': '通用預設', 'keywords': list(core.DEFAULT_WORDS)}
        if not self.settings.exists():
            return None, generic, ''
        try:
            settings = read_json(self.settings, 16 * 1024)
            if not isinstance(settings, dict) or type(settings.get('schema_version')) is not int or settings['schema_version'] != SCHEMA_VERSION or 'startup_profile' not in settings:
                raise ProfileError('不支援的啟動設定格式。')
            identifier = settings['startup_profile']
            return (None, generic, '') if identifier is None else (identifier, self.load(identifier), '')
        except ProfileError:
            return None, generic, '啟動設定或清單無法讀取，已使用通用預設；原設定檔保留，請重新載入清單。'
