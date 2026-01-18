"""
Модуль для сжатия изображений до указанного DPI.
"""
import logging
from pathlib import Path
from typing import List, Optional
from PIL import Image

logger = logging.getLogger(__name__)


class PhotoCompressor:
    """Сжимает изображения до указанного DPI перед вставкой в Word."""

    EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'}
    DEFAULT_DPI = 250

    def __init__(self, target_dpi: int = DEFAULT_DPI):
        """Инициализация компрессора.
        
        Args:
            target_dpi: Целевое разрешение DPI.
        """
        self.target_dpi = target_dpi

    def compress_image(self, input_path: Path, output_path: Optional[Path] = None) -> Path:
        """Сжимает фото до target_dpi.
        
        Args:
            input_path: Путь к исходному изображению.
            output_path: Путь для сохранения (если None, заменяет оригинал).
            
        Returns:
            Путь к сжатому файлу.
            
        Raises:
            FileNotFoundError: Если исходный файл не найден.
            IOError: При ошибке обработки изображения.
        """
        if not input_path.exists():
            raise FileNotFoundError(f"Файл не найден: {input_path}")
            
        if output_path is None:
            output_path = input_path

        try:
            with Image.open(input_path) as img:
                # Конвертируем в RGB, если PNG с альфой
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")

                # Пересохраняем с нужным DPI
                img.save(output_path, dpi=(self.target_dpi, self.target_dpi), quality=95)
                logger.debug(f"Сжато: {input_path} -> {output_path}")
                
        except Exception as e:
            logger.error(f"Ошибка при сжатии {input_path}: {e}")
            raise IOError(f"Не удалось сжать изображение: {e}")

        return output_path

    def compress_folder(self, folder: Path, inplace: bool = True) -> List[Path]:
        """Сжимает все изображения в папке и подпапках.
        
        Args:
            folder: Путь к папке с изображениями.
            inplace: Если True, заменяет файлы; иначе создаёт копии с префиксом.
            
        Returns:
            Список путей к сжатым файлам.
        """
        if not folder.exists():
            logger.warning(f"Папка не существует: {folder}")
            return []
            
        compressed_files = []
        for img_path in sorted(folder.rglob("*")):
            if img_path.suffix.lower() in self.EXTENSIONS and img_path.is_file():
                try:
                    out_path = img_path if inplace else folder / f"compressed_{img_path.name}"
                    self.compress_image(img_path, out_path)
                    compressed_files.append(out_path)
                except Exception as e:
                    logger.warning(f"Пропущен файл {img_path}: {e}")
                    
        logger.info(f"Сжато {len(compressed_files)} изображений в {folder}")
        return compressed_files
