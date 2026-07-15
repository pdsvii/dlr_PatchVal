#!/usr/bin/env python3
"""
Simple changelog generator for Patch Validator.

Usage:
  python scripts/generate_changelog.py --from-git
  python scripts/generate_changelog.py --release 0.1.0

- `--from-git` will collect commits since the latest tag and append them under `Unreleased`.
- `--release VERSION` will move the `Unreleased` section into a new release section with today's date.
"""

import argparse
import datetime
import subprocess
import os
import sys
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CHANGELOG_PATH = os.path.join(ROOT, 'CHANGELOG.md')


def load_changelog(path):
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()


def save_changelog(path, text):
    with open(path, 'w', encoding='utf-8') as f:
        f.write(text)


def find_unreleased_section(text):
    # Finds the Unreleased header and returns (start_idx, end_idx, content)
    m = re.search(r"## \[Unreleased\]\s*(.*?)\s*(?=\n## \[|\Z)", text, flags=re.S)
    if not m:
        return None
    return m.start(1), m.end(1), m.group(1).strip()


def release_version(version, changelog_path=CHANGELOG_PATH):
    text = load_changelog(changelog_path)
    sec = find_unreleased_section(text)
    if not sec:
        print('Unreleased section not found in', changelog_path)
        return 1
    start, end, content = sec
    if not content.strip():
        print('Unreleased section is empty; nothing to release.')
        return 1
    date = datetime.date.today().isoformat()
    release_header = f"## [{version}] - {date}\n\n"
    # Insert release_header + content after the Unreleased section
    before = text[:start]
    after = text[end:]
    # Build new Unreleased: keep the header but empty templates
    new_unreleased = '\n\n### Added\n- \n\n### Changed\n- \n\n### Fixed\n- \n'
    new_text = text[:text.find('## [Unreleased]') + len('## [Unreleased]')] + '\n\n' + new_unreleased + '\n' + release_header + content + '\n' + after
    save_changelog(changelog_path, new_text)
    print(f'Released {version} into {changelog_path}')
    return 0


def get_latest_tag():
    try:
        p = subprocess.run(['git', 'describe', '--tags', '--abbrev=0'], capture_output=True, text=True, check=True)
        return p.stdout.strip()
    except Exception:
        return None


def get_commits_since(tag=None):
    # Format: date|hash|subject
    fmt = '%ad|%h|%s'
    cmd = ['git', 'log', f'--pretty=format:{fmt}', '--date=short']
    if tag:
        cmd.insert(2, f'{tag}..HEAD')
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, check=True)
        lines = [l for l in p.stdout.splitlines() if l.strip()]
        commits = []
        for l in lines:
            parts = l.split('|', 2)
            if len(parts) == 3:
                commits.append({'date': parts[0], 'hash': parts[1], 'subject': parts[2]})
        return commits
    except subprocess.CalledProcessError as e:
        print('git command failed:', e)
        return []


def append_commits_to_unreleased(commits, changelog_path=CHANGELOG_PATH):
    if not commits:
        print('No commits to add.')
        return 1
    text = load_changelog(changelog_path)
    # Find Unreleased section insertion point: after its header
    m_header = re.search(r"(## \[Unreleased\]\s*)", text)
    if not m_header:
        print('Unreleased header not found in changelog.')
        return 1
    idx = m_header.end(1)
    lines = []
    for c in commits:
        lines.append(f"- {c['date']}: {c['subject']} ({c['hash']})")
    insertion = '\n'.join(lines) + '\n\n'
    new_text = text[:idx] + '\n' + insertion + text[idx:]
    save_changelog(changelog_path, new_text)
    print(f'Appended {len(commits)} commits to Unreleased in {changelog_path}')
    return 0


def main():
    parser = argparse.ArgumentParser(description='Changelog generator for Patch Validator')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--from-git', action='store_true', help='Collect commits since last tag and append to Unreleased')
    group.add_argument('--release', metavar='VERSION', help='Move Unreleased into a release section with VERSION')
    args = parser.parse_args()

    if not os.path.exists(CHANGELOG_PATH):
        print('CHANGELOG.md not found at', CHANGELOG_PATH)
        sys.exit(2)

    if args.from_git:
        tag = get_latest_tag()
        if tag:
            print('Latest tag found:', tag)
        else:
            print('No tags found; using all commits')
        commits = get_commits_since(tag)
        sys.exit(append_commits_to_unreleased(commits))

    if args.release:
        sys.exit(release_version(args.release))


if __name__ == '__main__':
    main()
