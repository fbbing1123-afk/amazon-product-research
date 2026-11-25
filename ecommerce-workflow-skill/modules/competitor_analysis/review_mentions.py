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


def report_points(analysis, kind):
    """区分没有入选差评和分析后没有具体原因，不能都用斜杠表示。"""
    if analysis[kind]:
        points = analysis[kind]
        if kind == "disadvantages":
            # 仅调整报告显示顺序，占比相同保持原顺序，不改变分析证据。
            points = sorted(points, key=lambda point: -mention_stats(
                point.get("review_numbers"), point.get("sample_count"))["mention_percent"])
        return [format_report_point(point, kind) for point in points]
    if kind == "disadvantages":
        count = analysis.get("selected_counts", {}).get(kind)
        if type(count) is not int or count < 0:
            raise ValueError("缺少已核验的入选差评数量，不能判断空结论含义")
        if count:
            return ["无具体差评原因"]
    return ["/"]
