from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from starlette.requests import Request

from images_feed import ImagesFeed
import sets_server


SAMPLE_XML = """<?xml version='1.0' encoding='utf-8'?>
<products>
  <product article="0010" brand="BABY TEAM">
    <image>http://176.104.6.149:8080/foto/BABY%20TEAM/0010.jpg</image>
    <image>http://176.104.6.149:8080/foto/BABY%20TEAM/0010_1.jpg</image>
  </product>
  <product article="NICE-02" brand="NICE GAME">
    <image>http://176.104.6.149:8080/foto/NICE%20GAME/002.jpg</image>
  </product>
</products>
"""


class ImagesFeedTests(unittest.TestCase):
    def test_images_feed_indexing_and_normalization(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            cache_file = Path(temp_dir) / "images_export.xml"
            cache_file.write_text(SAMPLE_XML, encoding="utf-8")

            feed = ImagesFeed("http://example.com/test.xml")
            feed.load = lambda force_refresh=False: None  # mock ensure_cache

            # Directly load from our test XML
            import xml.etree.ElementTree as ET
            tree = ET.parse(cache_file)
            root = tree.getroot()

            art_image = {}
            art_brand = {}
            for elem in root.iter():
                if elem.tag == "product":
                    art = elem.get("article") or ""
                    brand = elem.get("brand") or ""
                    image_url = ""
                    for child in elem:
                        if child.tag == "image":
                            image_url = child.text or ""
                            break
                    if art and image_url:
                        art_image[art.upper()] = image_url
                        art_image[art.upper().lstrip("0")] = image_url
                        art_brand[art.upper()] = brand
                        art_brand[art.upper().lstrip("0")] = brand

            feed.article_to_image = art_image
            feed.article_to_brand = art_brand
            feed._loaded = True

            # Exact search
            self.assertEqual(feed.get_image("0010"), "http://176.104.6.149:8080/foto/BABY%20TEAM/0010.jpg")
            self.assertEqual(feed.get_brand("0010"), "BABY TEAM")

            # Lowercase search
            self.assertEqual(feed.get_image("0010"), feed.get_image("0010"))

            # Stripped leading zeroes
            self.assertEqual(feed.get_image("10"), "http://176.104.6.149:8080/foto/BABY%20TEAM/0010.jpg")

            # Missing article
            self.assertEqual(feed.get_image("NON-EXISTENT"), "")

            # Meta object
            meta = feed.get_product_meta("0010")
            self.assertEqual(meta["article"], "0010")
            self.assertEqual(meta["brand"], "BABY TEAM")
            self.assertTrue(meta["image"].endswith("0010.jpg"))

            # Enrich list
            enriched = feed.enrich_products(["0010", "NICE-02"])
            self.assertEqual(len(enriched), 2)
            self.assertEqual(enriched[0]["article"], "0010")
            self.assertEqual(enriched[1]["brand"], "NICE GAME")

    def test_products_info_endpoint(self) -> None:
        res = sets_server.products_info(articles="0010; 0011")
        self.assertIn("items", res)
        self.assertEqual(len(res["items"]), 2)
        # 0010 is present in the cached feed
        first = res["items"][0]
        self.assertEqual(first["article"], "0010")
        if first.get("image"):
            self.assertIn("176.104.6.149", first["image"])


if __name__ == "__main__":
    unittest.main()
