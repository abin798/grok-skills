#!/usr/bin/env python3
"""Collect evidenced Amazon US organic results; no API credentials required.

Uses only Python's standard library. curl_cffi or requests is preferred when
already available. Offline fixtures never make network requests.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
from html.parser import HTMLParser
import json
from pathlib import Path
import re
import sys
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, urlencode, urljoin, urlparse
from urllib.request import HTTPRedirectHandler, Request, build_opener


ASIN = re.compile(r"^[A-Z0-9]{10}$")
SPACE = re.compile(r"\s+")
USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36"
VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"}


class Node:
    def __init__(self, tag="root", attrs=(), parent=None):
        self.tag, self.attrs, self.parent = tag, dict(attrs), parent
        self.children = []

    def text(self):
        return SPACE.sub(" ", " ".join(c.text() if isinstance(c, Node) else c for c in self.children)).strip()

    def nodes(self):
        for child in self.children:
            if isinstance(child, Node):
                yield child
                yield from child.nodes()

    def find(self, predicate):
        return next((n for n in self.nodes() if predicate(n)), None)

    def ident(self, value):
        return self.find(lambda n: n.attrs.get("id") == value)

    def klass(self, value):
        return self.find(lambda n: value in n.attrs.get("class", "").split())


class TreeParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node()
        self.current = self.root

    def handle_starttag(self, tag, attrs):
        node = Node(tag, attrs, self.current)
        self.current.children.append(node)
        if tag not in VOID:
            self.current = node

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag not in VOID:
            self.handle_endtag(tag)

    def handle_endtag(self, tag):
        candidate = self.current
        while candidate.parent:
            if candidate.tag == tag:
                self.current = candidate.parent
                return
            candidate = candidate.parent

    def handle_data(self, data):
        self.current.children.append(data)


def parse_html(text):
    parser = TreeParser()
    parser.feed(text)
    return parser.root


class FetchFailure(RuntimeError):
    pass


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


class Fetcher:
    def __init__(self, offline=None):
        self.offline = Path(offline).resolve() if offline else None
        self.backend = "offline" if offline else "urllib"
        self.session = None
        if not offline:
            try:
                from curl_cffi import requests
                self.session = requests.Session(impersonate="chrome")
                self.backend = "curl_cffi"
            except ImportError:
                try:
                    import requests
                    self.session = requests.Session()
                    self.backend = "requests"
                except ImportError:
                    pass
        self.opener = build_opener(NoRedirect())

    def get(self, url, fixture, binary=False):
        if self.offline:
            path = self.offline / fixture
            if not path.is_file():
                raise FetchFailure(f"离线文件缺失：{fixture}")
            content = path.read_bytes()
        else:
            try:
                if self.session:
                    response = self.session.get(url, timeout=25, allow_redirects=False,
                                                headers={"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"})
                    status, content = response.status_code, response.content
                else:
                    with self.opener.open(Request(url, headers={"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"}), timeout=25) as response:
                        status, content = response.status, response.read(20_000_001)
                if status != 200:
                    raise FetchFailure(f"HTTP {status}；停止，不重试或跟随重定向")
            except (HTTPError, URLError, OSError) as exc:
                raise FetchFailure(f"请求失败：{type(exc).__name__}；停止，不重试") from exc
            except Exception as exc:
                if isinstance(exc, FetchFailure):
                    raise
                raise FetchFailure(f"请求失败：{type(exc).__name__}；停止，不重试") from exc
        if len(content) > 20_000_000:
            raise FetchFailure("响应超过20MB，停止")
        if binary:
            return content
        text = content.decode("utf-8", errors="replace")
        if re.search(r'(?:id=["\']captchacharacters|<title>\s*(?:Amazon\s+)?(?:Robot Check|CAPTCHA)|Enter the characters you see below|Sorry, we just need to make sure you.re not a robot)', text, re.I):
            raise FetchFailure("页面出现验证码或机器人检查，停止；需要通过正常浏览器取得可访问页面")
        return text


def number(text, decimal=False):
    if not text:
        return None
    match = re.search(r"[\d,]+(?:\.\d+)?", text)
    if not match:
        return None
    value = float(match.group().replace(",", ""))
    return value if decimal else int(value)


def sponsored(node):
    current = node
    while current and current.tag != "root":
        if ("AdHolder" in current.attrs.get("class", "") or
                "s-sponsored" in current.attrs.get("data-component-type", "").lower()):
            return True
        current = current.parent
    for item in [node, *node.nodes()]:
        label = " ".join(item.attrs.get(k, "") for k in ("aria-label", "data-component-type", "data-cel-widget"))
        if re.search(r"\bsponsored\b", label, re.I):
            return True
        if item.tag in {"span", "a", "div"} and re.fullmatch(r"Sponsored(?:\s*[ⓘ?])?", item.text(), re.I):
            return True
    return False


def extract_price(node, original=False):
    candidates = [n for n in node.nodes() if "a-price" in n.attrs.get("class", "").split()]
    for candidate in candidates:
        is_original = "a-text-price" in candidate.attrs.get("class", "").split()
        if is_original == original:
            hidden = candidate.klass("a-offscreen")
            if hidden:
                return number(hidden.text(), decimal=True)
    return None


def observed_postcode(tree):
    line = tree.ident("glow-ingress-line2")
    if line:
        match = re.search(r"\b(\d{5})(?:-\d{4})?\b", line.text())
        if match:
            return match.group(1)
    return None


def extract_results(tree, page, source):
    result, position = [], 0
    for node in tree.nodes():
        if node.attrs.get("data-component-type") != "s-search-result":
            continue
        asin = node.attrs.get("data-asin", "").upper()
        if not ASIN.fullmatch(asin) or sponsored(node):
            continue
        position += 1
        record = {"页码": page, "页内位置": position, "ASIN": asin,
                  "_sources": {"ASIN": source, "页码": source, "页内位置": source}}
        heading = node.find(lambda n: n.tag == "h2")
        if heading:
            record["标题"] = heading.text()
        record["价格"] = extract_price(node)
        record["原价/参考价"] = extract_price(node, True)
        icon = node.klass("a-icon-alt")
        if icon and re.search(r"out of 5|5 stars", icon.text(), re.I):
            record["评分"] = number(icon.text(), True)
        for candidate in node.nodes():
            text = candidate.text()
            if ("s-underline-text" in candidate.attrs.get("class", "").split() and re.fullmatch(r"[\d,]+", text)):
                record["评论数"] = number(text)
                break
            label = candidate.attrs.get("aria-label", "")
            if re.fullmatch(r"[\d,]+ (?:ratings|reviews)", label, re.I):
                record["评论数"] = number(label)
                break
        bought = re.search(r"([\d,.]+)\s*([KM])?\s*(\+)?\s+bought in past month", node.text(), re.I)
        if bought:
            multiplier = {None: 1, "K": 1000, "M": 1000000}[bought.group(2).upper() if bought.group(2) else None]
            record["近一月销量"] = int(float(bought.group(1).replace(",", "")) * multiplier)
            record["_sales_lower_bound"] = bool(bought.group(3))
            record["_sales_display"] = bought.group(0)
            record["_sales_estimated"] = False
        image = node.find(lambda n: n.tag == "img" and "s-image" in n.attrs.get("class", "").split())
        if image:
            record["_image_url"] = image.attrs.get("src")
        for field in list(record):
            if not field.startswith("_") and record[field] is not None:
                record["_sources"].setdefault(field, source)
        result.append(record)
    return result


def detail_fields(tree, source):
    fields = {}
    for ident, key in (("productTitle", "标题"), ("acrCustomerReviewText", "评论数")):
        node = tree.ident(ident)
        if node:
            fields[key] = number(node.text()) if key == "评论数" else node.text()
    brand = tree.ident("bylineInfo")
    if brand:
        text = brand.text()
        if text.startswith("Brand:"):
            fields["品牌"] = text.partition(":")[2].strip()
        elif text.startswith("Visit the ") and text.endswith(" Store"):
            fields["品牌"] = text[len("Visit the "):-len(" Store")]
    price_root = tree.ident("corePriceDisplay_desktop_feature_div") or tree.ident("corePrice_feature_div")
    if price_root:
        fields["价格"] = extract_price(price_root)
        fields["原价/参考价"] = extract_price(price_root, True)
    availability = tree.ident("availability")
    if availability:
        text = availability.text()
        remaining = re.search(r"Only\s+(\d+)\s+left", text, re.I)
        if remaining:
            fields["库存状态"] = f"库存紧张（剩{remaining.group(1)}）"
        elif re.search(r"\bIn Stock\b", text, re.I):
            fields["库存状态"] = "有货"
        elif re.search(r"Currently unavailable", text, re.I) and not tree.ident("add-to-cart-button"):
            fields["库存状态"] = "无购买框"
    seller = tree.ident("sellerProfileTriggerId")
    if seller:
        seller_id = parse_qs(urlparse(urljoin(source, seller.attrs.get("href", ""))).query).get("seller", [None])[0]
        if seller_id and re.fullmatch(r"[A-Z0-9]+", seller_id):
            fields.update({"卖家ID": seller_id, "卖家": seller.text(), "店铺名称": seller.text(),
                           "卖家主页": f"https://www.amazon.com/sp?{urlencode({'seller': seller_id})}"})
    for row in tree.nodes():
        if row.tag not in {"tr", "li"}:
            continue
        text = row.text()
        if re.match(r"Date First Available\s*[:：]?", text, re.I):
            raw = re.sub(r"^Date First Available\s*[:：]?\s*", "", text, flags=re.I)
            for fmt in ("%B %d, %Y", "%b %d, %Y", "%Y-%m-%d"):
                try:
                    fields["上架时间"] = dt.datetime.strptime(raw, fmt).date().isoformat()
                    break
                except ValueError:
                    pass
        if re.match(r"Best Sellers Rank\s*[:：]?", text, re.I):
            ranks = re.findall(r"#\s*([\d,]+)\s+in\s+(.+?)(?=\s*\(|\s*#|$)", text)
            if ranks:
                fields.update({"大类BSR排名": number(ranks[0][0]), "大类": ranks[0][1].strip()})
            if len(ranks) > 1:
                fields.update({"小类BSR排名": number(ranks[1][0]), "小类名称": ranks[1][1].strip()})
    image = tree.ident("landingImage")
    if image:
        image_url = image.attrs.get("data-old-hires") or image.attrs.get("src")
        dynamic = image.attrs.get("data-a-dynamic-image")
        if not image_url and dynamic:
            try:
                image_url = next(iter(json.loads(dynamic)), None)
            except (ValueError, TypeError):
                pass
        if image_url:
            fields["_image_url"] = image_url
    return {key: value for key, value in fields.items() if value is not None}


def save_image(fetcher, record, directory):
    url = record.get("_image_url")
    if not url:
        return "无可核实的主图地址"
    parsed = urlparse(url)
    hostname = parsed.hostname or ""
    if parsed.scheme != "https" or not any(hostname == domain or hostname.endswith("." + domain) for domain in ("media-amazon.com", "ssl-images-amazon.com", "images-amazon.com")):
        return "主图不是可直接下载的Amazon HTTPS图片地址"
    fixture = "images/" + Path(parsed.path).name
    try:
        data = fetcher.get(url, fixture, binary=True)
    except FetchFailure as exc:
        return str(exc)
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        suffix = ".png"
    elif data.startswith(b"\xff\xd8\xff"):
        suffix = ".jpg"
    else:
        return "主图响应不是PNG或JPEG"
    destination = directory / (record["ASIN"] + suffix)
    destination.write_bytes(data)
    record["_image"] = "images/" + destination.name
    record["_sources"]["主图"] = url
    return None


def collect(args):
    work = Path(args.work_dir).resolve()
    evidence = work / "evidence"
    images = work / "images"
    evidence.mkdir(parents=True, exist_ok=True)
    images.mkdir(exist_ok=True)
    fetcher = Fetcher(args.offline_html_dir)
    captured = dt.datetime.now(dt.timezone.utc).isoformat()
    result = {"keyword": args.keyword, "marketplace": "US", "captured_at": captured,
              "delivery_postcode": None, "requested_postcode": args.postcode,
              "records": [], "variants": [], "notes": [], "points_used": 0,
              "status": "search_complete", "search_coverage_complete": False, "requested_pages": args.pages,
              "pages_collected": [], "fetch_backend": fetcher.backend}
    if args.offline_html_dir:
        result["notes"].append("离线HTML验证/导入；抓取时间记录为本次处理时间，原始页面时间须由调用方另行提供。")
    result["notes"].append("自然结果按页面DOM出现顺序计页内位置，剔除Sponsored文字/标签及广告组件后按ASIN去重；自然排名仅代表本次页面快照。")
    seen, signatures, postcodes = set(), set(), []
    reached_end = False
    for page in range(1, args.pages + 1):
        url = "https://www.amazon.com/s?" + urlencode({"k": args.keyword, "page": page})
        try:
            html = fetcher.get(url, f"page{page}.html")
            tree = parse_html(html)
            records = extract_results(tree, page, url)
            if not records:
                no_results = bool(re.search(r"No results for|No results found", tree.text(), re.I))
                if no_results:
                    result["notes"].append(f"第{page}页明确显示无结果，后续页不再请求。")
                    reached_end = True
                else:
                    raise FetchFailure("未解析到自然结果，也未见明确无结果提示；页面结构或访问状态待核实")
            signature = tuple(item["ASIN"] for item in records)
            if records and signature in signatures:
                raise FetchFailure("自然ASIN序列与先前页面完全相同，无法证明分页成功")
            signatures.add(signature)
            (evidence / f"page{page}.html").write_text(html, encoding="utf-8")
            result["pages_collected"].append({"page": page, "url": url, "sha256": hashlib.sha256(html.encode()).hexdigest(), "organic_items_before_dedup": len(records)})
            postcodes.append(observed_postcode(tree))
            for record in records:
                if record["ASIN"] in seen:
                    continue
                seen.add(record["ASIN"])
                record["自然排名"] = len(result["records"]) + 1
                record["_sources"]["自然排名"] = url
                result["records"].append(record)
            if reached_end:
                break
        except FetchFailure as exc:
            result["status"] = "incomplete"
            result["notes"].append(f"第{page}页未完成：{exc}。未取得页面不以数据库搜索结果补作自然排名。")
            break
    result["search_coverage_complete"] = reached_end or len(result["pages_collected"]) == args.pages
    if postcodes and postcodes[0] and all(value == postcodes[0] for value in postcodes):
        result["delivery_postcode"] = postcodes[0]
    else:
        result["notes"].append("配送邮编未能从所有已采页面一致核实，未声称设置成功。")
    if args.postcode and result["delivery_postcode"] != args.postcode:
        result["status"] = "incomplete"
        result["notes"].append(f"请求配送邮编{args.postcode}未被页面核实；本助手不自动修改Amazon地址会话。需在正常浏览器设置后导入页面。")
    detail_count = 0
    if args.details:
        for record in result["records"][:args.max_details]:
            asin = record["ASIN"]
            url = f"https://www.amazon.com/dp/{asin}"
            try:
                html = fetcher.get(url, f"details/{asin}.html")
                (evidence / f"detail_{asin}.html").write_text(html, encoding="utf-8")
                detail_tree = parse_html(html)
                detail_asin = detail_tree.ident("ASIN")
                if detail_asin and detail_asin.attrs.get("value", asin).upper() != asin:
                    raise FetchFailure("详情页ASIN与请求ASIN不一致")
                fields = detail_fields(detail_tree, url)
                if not fields:
                    raise FetchFailure("未解析到可核实的详情字段")
                record.update(fields)
                for field in fields:
                    if not field.startswith("_"):
                        record["_sources"][field] = url
                detail_count += 1
            except FetchFailure as exc:
                result["notes"].append(f"详情补充在{asin}停止：{exc}；后续商品保留搜索页字段。")
                break
        if detail_count < len(result["records"]):
            result["notes"].append(f"商品详情补充覆盖{detail_count}/{len(result['records'])}；缺失字段留空，需后续补充。")
    else:
        result["notes"].append("本次采集仅含搜索页已展示字段；详情、变体和卖家主体待后续补充。")
    result["details_collected"] = detail_count
    result["details_complete"] = args.details and detail_count == len(result["records"])
    image_failures = []
    for record in result["records"]:
        error = save_image(fetcher, record, images)
        if error:
            image_failures.append(f"{record['ASIN']}：{error}")
    if image_failures:
        result["notes"].append("主图缺失：" + "；".join(image_failures))
    result["images_complete"] = not image_failures
    result["enrichment_complete"] = False
    if any("_sales_display" in record for record in result["records"]):
        result["notes"].append("近一月销量中的Amazon bought in past month为页面购买量展示；带+的数值仅为下限，原始文本保留在各商品_sales_display中，不能视为精确销量。")
    result["notes"].append("本助手没有调用卖家精灵，点数使用为0；品牌/上架时间/父ASIN/销量估算、公司主体地址须按Skill继续补证据。")
    output = work / "competitors.json"
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"output": str(output), "status": result["status"], "search_coverage_complete": result["search_coverage_complete"], "products": len(result["records"]), "images": len(result["records"]) - len(image_failures)}, ensure_ascii=False))
    return 0 if result["status"] == "search_complete" else 2


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--keyword", required=True)
    parser.add_argument("--work-dir", required=True)
    parser.add_argument("--pages", type=int, choices=range(1, 4), default=3)
    parser.add_argument("--postcode", help="请求邮编；仅核验页面显示，不伪称自动设置")
    parser.add_argument("--offline-html-dir", help="page1.html等、details/ASIN.html与images/原文件名；禁止网络请求")
    parser.add_argument("--details", action="store_true", help="依次抓商品详情；遇访问限制停止补充")
    parser.add_argument("--max-details", type=int, default=200)
    args = parser.parse_args()
    if not args.keyword.strip():
        parser.error("搜索词不能为空")
    if args.postcode and not re.fullmatch(r"\d{5}", args.postcode):
        parser.error("邮编必须是美国五位邮编")
    if args.max_details < 1:
        parser.error("--max-details必须大于0")
    return collect(args)


if __name__ == "__main__":
    sys.exit(main())
