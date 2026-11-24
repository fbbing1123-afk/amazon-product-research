"""按入选差评的完整提及序号计算占比，仅供竞品分析表显示。"""

from decimal import Decimal, ROUND_HALF_UP


MENTION_SCOPE = "all_matching_selected_reviews"


def mention_stats(review_numbers, sample_count):
    """同一评论对同一原因只计一次，分母为本商品本次全部入选差评。"""
    if (type(sample_count) is not int or sample_count <= 0
            or not isinstance(review_numbers, list) or not review_numbers
            or any(type(n) is not int or not 1 <= n <= sample_count for n in review_numbers)
            or len(set(review_numbers)) != len(review_numbers)):
        raise ValueError("提及占比的样本数或评论序号无效")
    count = len(review_numbers)
    percent = (Decimal(count) * 100 / Decimal(sample_count)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
    return {"mention_count": count, "sample_count": sample_count, "mention_percent": float(percent)}


def format_report_point(point, kind):
    """仅在竞品分析表的缺点后附加百分比，不显示提及条数或样本数。"""
    if kind != "disadvantages":
        return point["text"]
    stats = mention_stats(point.get("review_numbers"), point.get("sample_count"))
    if any(point.get(key) != value for key, value in stats.items()):
        raise ValueError("缺少有效提及统计，请按新提示重新分析并校验")
    return f"{point['text']}（{stats['mention_percent']:g}%）"
