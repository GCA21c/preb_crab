from __future__ import annotations

import json
import zipfile
from pathlib import Path

from PySide6.QtGui import QImage

from core.clipboard_store import ClipboardItem, ClipboardStore


class ProjectStore:
    def save(
        self,
        output_path: str,
        clipboard_store: ClipboardStore,
        here_pages: list[list[dict]],
        drawing_pages: list[list[dict]] | None = None,
    ) -> str:
        out = Path(output_path)
        if out.suffix.lower() != '.dcap':
            out = out.with_suffix('.dcap')
        manifest = {
            'version': 3,
            'clipboard': [],
            'here_pages': [],
            'drawing_pages': drawing_pages or [[] for _ in here_pages],
        }
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
            for idx, item in enumerate(clipboard_store.items, start=1):
                name = f'clipboard/item_{idx:04d}.png'
                item.image.save(str(out.parent / f'__tmp_clip_{idx:04d}.png'))
                tmp = out.parent / f'__tmp_clip_{idx:04d}.png'
                zf.write(tmp, name)
                tmp.unlink(missing_ok=True)
                manifest['clipboard'].append({
                    'number': item.number,
                    'timestamp': item.timestamp,
                    'name': getattr(item, 'name', item.timestamp),
                    'image': name,
                })

            for page_idx, page in enumerate(here_pages, start=1):
                page_manifest = []
                for block_idx, block in enumerate(page, start=1):
                    name = f'here/page_{page_idx:03d}_block_{block_idx:03d}.png'
                    block['image'].save(str(out.parent / f'__tmp_page_{page_idx:03d}_block_{block_idx:03d}.png'))
                    tmp = out.parent / f'__tmp_page_{page_idx:03d}_block_{block_idx:03d}.png'
                    zf.write(tmp, name)
                    tmp.unlink(missing_ok=True)
                    page_manifest.append({
                        'image': name,
                        'x': block['x'],
                        'y': block['y'],
                        'w': block['w'],
                        'h': block['h'],
                        'original_w': block.get('original_w', block['w']),
                        'original_h': block.get('original_h', block['h']),
                        'source_index': block.get('source_index', -1),
                        'content_left': block.get('content_left', 0),
                        'content_right': block.get('content_right', block.get('w', 0)),
                    })
                manifest['here_pages'].append(page_manifest)

            zf.writestr('manifest.json', json.dumps(manifest, ensure_ascii=False, indent=2))
        return str(out)

    def load(self, input_path: str) -> dict:
        src = Path(input_path)
        with zipfile.ZipFile(src, 'r') as zf:
            manifest = json.loads(zf.read('manifest.json').decode('utf-8'))
            clipboard_items: list[ClipboardItem] = []
            for entry in manifest.get('clipboard', []):
                data = zf.read(entry['image'])
                image = QImage.fromData(data)
                clipboard_items.append(ClipboardItem(
                    number=entry['number'],
                    timestamp=entry['timestamp'],
                    name=entry.get('name', entry['timestamp']),
                    image=image,
                ))

            pages: list[list[dict]] = []
            for page in manifest.get('here_pages', []):
                page_blocks = []
                for block in page:
                    data = zf.read(block['image'])
                    image = QImage.fromData(data)
                    page_blocks.append({
                        'image': image,
                        'x': block['x'],
                        'y': block['y'],
                        'w': block['w'],
                        'h': block['h'],
                        'original_w': block.get('original_w', block['w']),
                        'original_h': block.get('original_h', block['h']),
                        'source_index': block.get('source_index', -1),
                        'content_left': block.get('content_left', 0),
                        'content_right': block.get('content_right', block.get('w', 0)),
                    })
                pages.append(page_blocks)
        drawing_pages = manifest.get('drawing_pages') or [[] for _ in pages]
        return {'clipboard_items': clipboard_items, 'here_pages': pages or [[]], 'drawing_pages': drawing_pages}
