#!/usr/bin/env python3
"""
IOMS eHub -> SharePoint Migration Parser

Parses IOMS HTML files and produces:
  1. migration_manifest.csv  - one row per page, all SharePoint metadata
  2. image_manifest.json     - one entry per image, with associated screen codes

Usage:
    python parse_ioms.py <html_folder> [output.csv] [images.json]

Requirements:
    pip install beautifulsoup4
"""

import os
import re
import csv
import json
import sys
from pathlib import Path
from bs4 import BeautifulSoup

# Matches IOMS screen/window codes like CSP440, CP110, VR100, CSP3, etc.
SCREEN_CODE_RE = re.compile(r'\(([A-Z]{2,6}\d{2,4})\)')


def get_hidden(soup, field_id: str) -> str:
    el = soup.find('input', {'id': field_id})
    return el['value'].strip() if el else ''


def detect_pagetype(soup) -> str:
    subheadings = [p.text.strip() for p in soup.find_all('p', class_='subheading2')]
    if any('Practice Steps' in s for s in subheadings):
        return 'Procedure'
    if any('Overview' in s for s in subheadings):
        return 'Overview'
    if any('Report' in s for s in subheadings):
        return 'Report'
    if soup.find('ol', class_='listnumber'):
        return 'Procedure'
    return 'Reference'


def extract_images(soup) -> list:
    """
    Extract all images and associate them with screen codes found in the
    nearest containing list item or paragraph.
    """
    images = []
    for img in soup.find_all('img'):
        src = img.get('src', '')
        # Image IDs are prefixed with 'f' (e.g. f16224 -> 16224)
        raw_id = img.get('id', '')
        image_id = raw_id.lstrip('f') if raw_id.startswith('f') else raw_id

        # Collect screen codes from the nearest containing block
        nearby_codes = []
        for ancestor in img.parents:
            tag = ancestor.name
            if tag in ('li', 'p', 'td', 'div', 'body'):
                codes = SCREEN_CODE_RE.findall(ancestor.get_text())
                nearby_codes.extend(codes)
                if tag in ('li', 'p', 'td'):
                    break  # stop at first meaningful block

        images.append({
            'src': src,
            'image_id': image_id,
            'screen_codes': list(dict.fromkeys(nearby_codes)),  # ordered dedup
        })
    return images


def parse_html_file(filepath: Path) -> dict:
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')

    # --- Core metadata ---
    topic_id = get_hidden(soup, 'topicId') or filepath.stem
    title = (
        get_hidden(soup, 'topicDescription')
        or (soup.title.string.strip() if soup.title else filepath.stem)
    )
    modified = get_hidden(soup, 'footer-modified').replace('Last modified: ', '').strip()
    modified_by = get_hidden(soup, 'footer-modifiedby').replace('Modified by: ', '').strip()
    legacy_file = f"#{topic_id}.htm"

    # --- Business area from top breadcrumb ---
    breadcrumb = soup.find('table', class_='relatedtopics aboveheading')
    business_area = ''
    if breadcrumb:
        link = breadcrumb.find('a')
        if link:
            business_area = link.text.strip()

    # --- Content type hint (e.g. "Probation & Parole") ---
    notepp = soup.find('p', class_='notepp')
    content_type = notepp.text.strip() if notepp else ''

    # --- Page type ---
    pagetype = detect_pagetype(soup)

    # --- Screen codes from full body text ---
    body = soup.find('body')
    body_text = body.get_text() if body else soup.get_text()
    screen_codes = list(dict.fromkeys(SCREEN_CODE_RE.findall(body_text)))

    # --- Images ---
    images = extract_images(soup)

    # --- Related links (below content) ---
    related_table = soup.find('table', class_='relatedtopics belowtopictext')
    related_links = []
    if related_table:
        for a in related_table.find_all('a'):
            href = a.get('href', '')
            related_links.append(f"{a.text.strip()}:{href}")

    return {
        'title': title,
        'page_key': topic_id,
        'legacy_file': legacy_file,
        'business_area': business_area,
        'content_type': content_type,
        'pagetype': pagetype,
        'screen_codes': screen_codes,
        'images': images,
        'related_links': related_links,
        'last_modified': modified,
        'modified_by': modified_by,
        'source_file': filepath.name,
    }


def process_folder(input_folder: str, csv_out: str, img_out: str):
    input_path = Path(input_folder)
    htm_files = sorted(
        list(input_path.glob('*.htm')) + list(input_path.glob('*.html'))
    )

    if not htm_files:
        print(f"No .htm/.html files found in: {input_folder}")
        sys.exit(1)

    print(f"Found {len(htm_files)} HTML files in {input_folder}\n")

    pages = []
    all_images = []
    errors = []

    for filepath in htm_files:
        try:
            data = parse_html_file(filepath)

            pages.append({
                'Title': data['title'],
                'PageKey': data['page_key'],
                'LegacyFile': data['legacy_file'],
                'BusinessArea': data['business_area'],
                'ContentType': data['content_type'],
                'Pagetype': data['pagetype'],
                'ScreenCodes': '|'.join(data['screen_codes']),
                'RelatedLinks': '|'.join(data['related_links']),
                'ImageCount': len(data['images']),
                'LastModified': data['last_modified'],
                'ModifiedBy': data['modified_by'],
                'Status': 'To Do',
                'SourceFile': data['source_file'],
            })

            for img in data['images']:
                all_images.append({
                    'PageKey': data['page_key'],
                    'PageTitle': data['title'],
                    'ImageSrc': img['src'],
                    'ImageID': img['image_id'],
                    'ScreenCodes': '|'.join(img['screen_codes']),
                })

            codes_str = ', '.join(data['screen_codes']) if data['screen_codes'] else 'none'
            print(f"  OK  {filepath.name:20s}  '{data['title'][:50]}'  codes=[{codes_str}]")

        except Exception as e:
            errors.append({'file': filepath.name, 'error': str(e)})
            print(f"  ERR {filepath.name}: {e}")

    # Write page manifest CSV
    if pages:
        with open(csv_out, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=pages[0].keys())
            writer.writeheader()
            writer.writerows(pages)
        print(f"\nPage manifest:  {csv_out}  ({len(pages)} pages)")

    # Write image manifest JSON
    with open(img_out, 'w', encoding='utf-8') as f:
        json.dump(all_images, f, indent=2, ensure_ascii=False)
    print(f"Image manifest: {img_out}  ({len(all_images)} images)")

    if errors:
        print(f"\nErrors ({len(errors)}):")
        for e in errors:
            print(f"  {e['file']}: {e['error']}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    folder = sys.argv[1]
    csv_path = sys.argv[2] if len(sys.argv) > 2 else 'migration_manifest.csv'
    img_path = sys.argv[3] if len(sys.argv) > 3 else 'image_manifest.json'

    process_folder(folder, csv_path, img_path)
