"""
Модуль завантаження та кешування зображень товарів з XML-фіду.
Використовується для візуального відображення товарів у наборах.
URL за замовчуванням: http://176.104.6.149:8080/xml/images_export.xml
"""
from __future__ import annotations

import json
import logging
import os
import re
import shutil
import threading
import time
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

import requests

logger = logging.getLogger("images_feed")

PROJECT_ROOT = Path(__file__).resolve().parent
DATA_DIR = PROJECT_ROOT / "data"
CACHE_DIR = DATA_DIR / "cache"
CACHE_FILE = CACHE_DIR / "images_export.xml"
CACHE_TTL_SECONDS = 3600 * 4  # 4 години кешування XML-фіду

# Сусідні папки, звідки можна миттєво підтягнути вже наявний локальний кеш
SEED_CACHE_LOCATIONS = [
    Path(r"C:\Users\ABM\PyCharmProject\Categorizer\data\cache\images_export.xml"),
    Path(r"C:\Users\ABM\PyCharmProject\changedetectionio\price_tracker\data\cache\images_export.xml"),
    Path(r"C:\Users\ABM\PyCharmProject\DockerPriceTrackingScript\data\cache\images_export.xml"),
]


def get_configured_xml_url(config: dict[str, Any] | None = None) -> str:
    if config and config.get("images_xml_url"):
        return str(config["images_xml_url"]).strip()

    cfg_file = PROJECT_ROOT / "config.json"
    if cfg_file.exists():
        try:
            with open(cfg_file, "r", encoding="utf-8-sig") as f:
                cfg = json.load(f)
                url = cfg.get("images_xml_url")
                if url:
                    return str(url).strip()
        except Exception:
            pass

    return "http://176.104.6.149:8080/xml/images_export.xml"


def _clean_code(val: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", val).upper()


class ImagesFeed:
    _instance: "ImagesFeed | None" = None
    _lock = threading.Lock()

    def __init__(self, xml_url: str | None = None) -> None:
        self.xml_url = xml_url or get_configured_xml_url()
        self.article_to_image: dict[str, str] = {}
        self.article_to_brand: dict[str, str] = {}
        self.article_to_name: dict[str, str] = {}
        self._loaded = False
        self._load_lock = threading.Lock()
        self._last_error = ""

    @classmethod
    def get_instance(cls, xml_url: str | None = None) -> "ImagesFeed":
        with cls._lock:
            if cls._instance is None:
                cls._instance = cls(xml_url)
            return cls._instance

    def _seed_from_existing(self) -> bool:
        if CACHE_FILE.exists() and CACHE_FILE.stat().st_size > 1024:
            return True
        for seed_path in SEED_CACHE_LOCATIONS:
            if seed_path.exists() and seed_path.stat().st_size > 1024:
                try:
                    CACHE_DIR.mkdir(parents=True, exist_ok=True)
                    shutil.copyfile(seed_path, CACHE_FILE)
                    logger.info("Скопійовано початковий кеш зображень з %s", seed_path)
                    return True
                except Exception as exc:
                    logger.warning("Не вдалося скопіювати кеш з %s: %s", seed_path, exc)
        return False

    def ensure_cache(self, force_refresh: bool = False) -> str | None:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        need_download = force_refresh or not CACHE_FILE.exists()

        if not need_download:
            try:
                mtime = CACHE_FILE.stat().st_mtime
                if time.time() - mtime > CACHE_TTL_SECONDS:
                    need_download = True
            except OSError:
                need_download = True

        if need_download:
            # Спочатку перевіримо, чи можна підтягнути свіжий кеш із сусіднього проекту
            if not force_refresh and self._seed_from_existing():
                need_download = False

        if need_download:
            logger.info("Завантаження XML-фіду зображень з %s...", self.xml_url)
            try:
                resp = requests.get(self.xml_url, timeout=35)
                resp.raise_for_status()
                with open(CACHE_FILE, "wb") as f:
                    f.write(resp.content)
                logger.info("XML-фід зображень успішно збережено в кеш (%d байт)", len(resp.content))
                self._last_error = ""
            except Exception as exc:
                self._last_error = str(exc)
                logger.warning("Не вдалося завантажити XML-фід зображень: %s", exc)
                if not CACHE_FILE.exists() and not self._seed_from_existing():
                    return None
                logger.info("Використовується наявний локальний кеш XML-фіду.")

        return str(CACHE_FILE) if CACHE_FILE.exists() else None

    def load(self, force_refresh: bool = False) -> None:
        with self._load_lock:
            if self._loaded and not force_refresh:
                return

            cache_path = self.ensure_cache(force_refresh=force_refresh)
            if not cache_path:
                logger.warning("Кеш зображень недоступний.")
                self._loaded = True
                return

            try:
                tree = ET.parse(cache_path)
                root = tree.getroot()

                art_image: dict[str, str] = {}
                art_brand: dict[str, str] = {}
                art_name: dict[str, str] = {}

                for elem in root.iter():
                    if elem.tag in ("product", "offer", "item"):
                        art = elem.get("article") or elem.get("id") or ""
                        brand = elem.get("brand") or ""
                        name = ""
                        image_url = ""

                        for child in elem:
                            tag_lower = child.tag.lower()
                            if not art and tag_lower in ("article", "articul", "vendorcode", "code"):
                                art = (child.text or "").strip()
                            if not brand and tag_lower in ("brand", "vendor"):
                                brand = (child.text or "").strip()
                            if not name and tag_lower in ("name", "title"):
                                name = (child.text or "").strip()
                            if not image_url and tag_lower in ("image", "picture", "photo"):
                                image_url = (child.text or "").strip()

                        art_raw = (art or "").strip()
                        if not art_raw:
                            continue

                        art_upper = art_raw.upper()
                        art_norm = _clean_code(art_upper)
                        art_lstrip = art_upper.lstrip("0")

                        keys_to_index = {art_upper}
                        if art_norm:
                            keys_to_index.add(art_norm)
                        if art_lstrip:
                            keys_to_index.add(art_lstrip)

                        if image_url:
                            for key in keys_to_index:
                                if key not in art_image:
                                    art_image[key] = image_url
                        if brand:
                            for key in keys_to_index:
                                if key not in art_brand:
                                    art_brand[key] = brand
                        if name:
                            for key in keys_to_index:
                                if key not in art_name:
                                    art_name[key] = name

                self.article_to_image = art_image
                self.article_to_brand = art_brand
                self.article_to_name = art_name
                self._loaded = True
                logger.info("Індексовано зображень для %d артикулів.", len(self.article_to_image))
            except Exception as exc:
                self._last_error = str(exc)
                logger.error("Помилка під час парсингу XML-фіду: %s", exc)
                self._loaded = True

    def is_loaded(self) -> bool:
        return self._loaded

    def get_image(self, article: str) -> str:
        if not self._loaded:
            self.load()
        if not article:
            return ""
        art_clean = str(article).strip().upper()
        if not art_clean:
            return ""

        # 1. Точний пошук
        url = self.article_to_image.get(art_clean)
        if url:
            return url

        # 2. Пошук без початкових нулів
        lstripped = art_clean.lstrip("0")
        if lstripped and lstripped in self.article_to_image:
            return self.article_to_image[lstripped]

        # 3. Нормалізований пошук (лише букви та цифри)
        norm = _clean_code(art_clean)
        if norm and norm in self.article_to_image:
            return self.article_to_image[norm]

        return ""

    def get_brand(self, article: str) -> str:
        if not self._loaded:
            self.load()
        if not article:
            return ""
        art_clean = str(article).strip().upper()
        if not art_clean:
            return ""
        if art_clean in self.article_to_brand:
            return self.article_to_brand[art_clean]
        lstripped = art_clean.lstrip("0")
        if lstripped in self.article_to_brand:
            return self.article_to_brand[lstripped]
        norm = _clean_code(art_clean)
        if norm in self.article_to_brand:
            return self.article_to_brand[norm]
        return ""

    def get_product_meta(self, article: str) -> dict[str, Any]:
        art_str = str(article or "").strip()
        return {
            "article": art_str,
            "image": self.get_image(art_str),
            "brand": self.get_brand(art_str),
        }

    def enrich_products(self, articles: list[str]) -> list[dict[str, Any]]:
        return [self.get_product_meta(art) for art in articles if str(art).strip()]


images_feed = ImagesFeed.get_instance()
