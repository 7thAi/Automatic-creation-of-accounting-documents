"""
Модуль для анализа папок с фотографиями.
"""
import logging
from pathlib import Path
from typing import List

logger = logging.getLogger(__name__)


class PhotoFolderAnalyzer:
    """Анализирует папку и подсчитывает количество фотографий."""

    EXTENSIONS: set = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'}

    def count_photos_in_folder(self, folder_path: Path) -> int:
        """Возвращает количество фото в папке и подпапках.
        
        Args:
            folder_path: Путь к папке для анализа.
            
        Returns:
            Количество найденных фотографий.
        """
        if not folder_path.exists():
            logger.debug(f"Папка не существует: {folder_path}")
            return 0
            
        count = sum(
            1 for f in folder_path.rglob("*")
            if f.is_file() and f.suffix.lower() in self.EXTENSIONS
        )
        logger.debug(f"Найдено {count} фото в {folder_path}")
        return count

    def get_photo_list(self, folder_path: Path) -> List[Path]:
        """Возвращает список путей к фотографиям в папке.
        
        Args:
            folder_path: Путь к папке для анализа.
            
        Returns:
            Список путей к фотографиям.
        """
        if not folder_path.exists():
            return []
            
        return sorted([
            f for f in folder_path.rglob("*")
            if f.is_file() and f.suffix.lower() in self.EXTENSIONS
        ])
