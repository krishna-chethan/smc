def num_to_words_inr(num):
    """
    Converts a number into Indian Rupee words.
    Example: 15300 -> Fifteen Thousand Three Hundred rupees only
    """
    try:
        num = float(num)
    except (ValueError, TypeError):
        return "Zero rupees only"
    
    if num == 0:
        return "Zero rupees only"

    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
            "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
            "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

    def two_digits(n):
        if n == 0:
            return ""
        if n < 20:
            return ones[n]
        return tens[n // 10] + (" " + ones[n % 10] if n % 10 != 0 else "")

    def three_digits(n):
        h = n // 100
        rem = n % 100
        res = ""
        if h > 0:
            res += ones[h] + " Hundred"
            if rem > 0:
                res += " and "
        if rem > 0:
            res += two_digits(rem)
        return res

    integer_part = int(abs(num))
    paisa_part = int(round((abs(num) - integer_part) * 100))

    crore = integer_part // 10000000
    integer_part %= 10000000
    lakh = integer_part // 100000
    integer_part %= 100000
    thousand = integer_part // 1000
    integer_part %= 1000
    hundreds = integer_part

    parts = []
    if crore > 0:
        parts.append(two_digits(crore) + " Crore")
    if lakh > 0:
        parts.append(two_digits(lakh) + " Lakh")
    if thousand > 0:
        parts.append(two_digits(thousand) + " Thousand")
    if hundreds > 0:
        parts.append(three_digits(hundreds))

    words = " ".join(filter(None, parts)).strip()
    if not words:
        words = "Zero"

    res = words + " rupees"
    if paisa_part > 0:
        res += " and " + two_digits(paisa_part) + " paise"
    res += " only"
    return res
