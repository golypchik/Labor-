#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Удаление последнего пустого листа альбомной ориентации из шаблонов."""
import sys
from pathlib import Path

# Добавляем корень проекта в путь
project_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(project_root))

from docx import Document
from docx.oxml.ns import qn


def remove_last_landscape_section(docx_path):
    """Удаляет последнюю секцию, если она альбомная и пустая."""
    doc = Document(docx_path)
    body = doc.element.body
    children = list(body)
    if len(children) < 2:
        return False
    # Последний элемент — обычно sectPr (свойства секции)
    last = children[-1]
    tag = last.tag.split('}')[-1] if '}' in last.tag else last.tag
    if tag == 'sectPr':
        # Проверяем, альбомная ли ориентация
        pgSz = last.find(qn('w:pgSz'))
        if pgSz is not None:
            Orient = pgSz.get(qn('w:orient'))
            if Orient == 'landscape':
                body.remove(last)
                doc.save(docx_path)
                return True
    return False


if __name__ == '__main__':
    temp_dir = project_root / 'temp'
    for name in ['template3.docx', 'template4.docx', 'template5.docx']:
        path = temp_dir / name
        if path.exists():
            try:
                if remove_last_landscape_section(path):
                    print(f'Удалена последняя альбомная секция из {name}')
                else:
                    print(f'Не найдена альбомная секция для удаления в {name}')
            except Exception as e:
                print(f'Ошибка при обработке {name}: {e}')
