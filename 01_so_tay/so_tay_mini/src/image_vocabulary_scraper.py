import json
import re
from pathlib import Path
from urllib.parse import urljoin, urlparse


LANGOAL_LESSON_URL_TEMPLATE = "https://langoal.com/vocbs/lesson-{lesson}"
IMAGE_URL_PATTERN_TEMPLATE = r"mnn-l{lesson}-(\d+)\.png$"
REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0 Safari/537.36"
    )
}


def format_lesson_dir_name(lesson):
    return f"lesson_{int(lesson):02d}"


def get_langoal_lesson_url(lesson):
    return LANGOAL_LESSON_URL_TEMPLATE.format(lesson=int(lesson))


def clean_text(text):
    return re.sub(r"\s+", " ", text or "").strip()


def pick_original_image_url(img, page_url):
    candidates = []
    for attr in ["src", "data-src"]:
        value = img.get(attr)
        if value:
            candidates.append(value)

    srcset = img.get("srcset")
    if srcset:
        for part in srcset.split(","):
            candidate = part.strip().split(" ", 1)[0]
            if candidate:
                candidates.append(candidate)

    absolute_candidates = [urljoin(page_url, candidate) for candidate in candidates]
    for candidate in absolute_candidates:
        if "-p-" not in Path(urlparse(candidate).path).name:
            return candidate
    return absolute_candidates[0] if absolute_candidates else ""


def get_nearest_title(img):
    title = clean_text(img.get("alt"))
    if title:
        return title

    card = img.find_parent("div", class_=lambda value: value and "items-center" in value)
    if card:
        heading = card.find(["h3", "h2"])
        if heading:
            return clean_text(heading.get_text(" "))
    return ""


def get_intro_image_container(soup):
    heading = soup.find(
        lambda tag: tag.name in {"h1", "h2", "h3"}
        and "新出語彙導入イラスト" in tag.get_text(" ", strip=True)
    )
    if not heading:
        return soup

    container = heading.find_next(
        lambda tag: tag.name == "div"
        and "category-grid" in (tag.get("class") or [])
    )
    return container or soup


def parse_langoal_vocabulary_items(html, lesson, page_url=None):
    from bs4 import BeautifulSoup

    lesson = int(lesson)
    page_url = page_url or get_langoal_lesson_url(lesson)
    image_pattern = re.compile(IMAGE_URL_PATTERN_TEMPLATE.format(lesson=lesson))
    soup = BeautifulSoup(html, "html.parser")
    container = get_intro_image_container(soup)

    items_by_index = {}
    for img in container.find_all("img"):
        image_url = pick_original_image_url(img, page_url)
        image_name = Path(urlparse(image_url).path).name
        match = image_pattern.search(image_name)
        if not match:
            continue

        index = int(match.group(1))
        vocabulary = get_nearest_title(img)
        if not vocabulary:
            continue

        items_by_index[index] = {
            "index": index,
            "vocabulary": vocabulary,
            "image_url": image_url,
        }

    return [items_by_index[index] for index in sorted(items_by_index)]


def create_webdriver(browser="auto", headless=True):
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.edge.options import Options as EdgeOptions

    browser_order = ["chrome", "edge"] if browser == "auto" else [browser]
    errors = []

    for browser_name in browser_order:
        try:
            if browser_name == "chrome":
                options = ChromeOptions()
                if headless:
                    options.add_argument("--headless=new")
                options.add_argument("--disable-gpu")
                options.add_argument("--window-size=1440,1600")
                return webdriver.Chrome(options=options)

            if browser_name == "edge":
                options = EdgeOptions()
                if headless:
                    options.add_argument("--headless=new")
                options.add_argument("--disable-gpu")
                options.add_argument("--window-size=1440,1600")
                return webdriver.Edge(options=options)
        except Exception as error:
            errors.append(f"{browser_name}: {error}")

    raise RuntimeError(
        "Không mở được Chrome/Edge bằng Selenium. "
        "Hãy cài Google Chrome hoặc Microsoft Edge bản mới, rồi chạy lại.\n"
        + "\n".join(errors)
    )


def fetch_langoal_lesson_html(lesson, browser="auto", headless=True, timeout=25):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait

    lesson = int(lesson)
    url = get_langoal_lesson_url(lesson)
    driver = create_webdriver(browser=browser, headless=headless)
    try:
        driver.get(url)
        selector = f"img[src*='mnn-l{lesson}-'], img[srcset*='mnn-l{lesson}-']"
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, selector))
        )
        return driver.page_source
    finally:
        driver.quit()


def download_image(image_url, destination_path, timeout=30):
    import requests

    destination_path.parent.mkdir(parents=True, exist_ok=True)
    if destination_path.exists() and destination_path.stat().st_size > 0:
        return False

    response = requests.get(
        image_url,
        headers=REQUEST_HEADERS,
        timeout=timeout,
    )
    response.raise_for_status()
    destination_path.write_bytes(response.content)
    return True


def save_metadata(metadata, data_dir):
    lesson_dir_name = format_lesson_dir_name(metadata["lesson"])
    output_dir = data_dir / "image_vocabulary"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{lesson_dir_name}.json"
    output_path.write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return output_path


def path_for_metadata(path, project_root):
    try:
        return path.resolve().relative_to(project_root.resolve()).as_posix()
    except ValueError:
        return path.as_posix()


def download_langoal_vocabulary_lesson(
    lesson,
    assets_dir,
    data_dir,
    project_root,
    browser="auto",
    headless=True,
    html=None,
):
    lesson = int(lesson)
    source_url = get_langoal_lesson_url(lesson)
    page_html = html if html is not None else fetch_langoal_lesson_html(
        lesson,
        browser=browser,
        headless=headless,
    )
    items = parse_langoal_vocabulary_items(page_html, lesson, page_url=source_url)
    if not items:
        raise RuntimeError(f"Không tìm thấy ảnh từ vựng ở {source_url}")

    lesson_dir_name = format_lesson_dir_name(lesson)
    image_dir = assets_dir / "images" / "vocabulary" / lesson_dir_name
    for item in items:
        image_path = image_dir / f"{lesson_dir_name}_{item['index']:02d}.png"
        download_image(item["image_url"], image_path)
        item["image_file"] = path_for_metadata(image_path, project_root)

    metadata = {
        "lesson": lesson,
        "source_url": source_url,
        "image_dir": path_for_metadata(image_dir, project_root),
        "items": items,
    }
    metadata_path = save_metadata(metadata, data_dir)
    return metadata, metadata_path
