from pprint import pprint

STANDARD_ENCODING = {
    "תשרי": 7,
    "חשוון": 8,
    "כסלו": 9,
    "טבת": 10,
    "שבט": 11,
    "אדר": 12,
    "אדרב": 13,
    "ניסן": 1,
    "אייר": 2,
    "סיוון": 3,
    "תמוז": 4,
    "אב": 5,
    "אלול": 6
}

inv_STANDARD_ENCODING = {v: k for k, v in STANDARD_ENCODING.items()}


class MonthConvertor:
    def decode_word(word: str) -> int:
        word = word.replace(u'\xa0', u'')
        if word.strip() in STANDARD_ENCODING:
            return STANDARD_ENCODING[word]
        else:
            print("the month : " + str(word) + "is not valid month"
                                               "enter alternate month number:")
            pprint(STANDARD_ENCODING)
            alternate_month = int(input())
            print(decode_num(alternate_month)+ "הוזן בהצלחה")
            return alternate_month


def decode_num(user_num):
    if int(user_num) < 14:
        return inv_STANDARD_ENCODING[user_num]
    else:
        print("חודש יכול להיות מספר בין 1 ל13 לפי הטבלה אנא נסה להזין בשנית את המספר של החודש הנכון")
        print("enter alternate month number:")
        pprint(STANDARD_ENCODING)
        alternate_month = int(input())
        print(decode_num(alternate_month))
        return alternate_month
