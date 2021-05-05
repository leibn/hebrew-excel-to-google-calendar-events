STANDARD_ENCODING = {
    "א": 1,
    "ב": 2,
    "ג": 3,
    "ד": 4,
    "ה": 5,
    "ו": 6,
    "ז": 7,
    "ח": 8,
    "ט": 9,
    "י": 10,
    "כ": 20,
    "ל": 30,
    "מ": 40,
    "נ": 50,
    "ס": 60,
    "ע": 70,
    "פ": 80,
    "צ": 90,
    "ק": 100,
    "ר": 200,
    "ש": 300,
    "ת": 400,
    "'": 0,
    " ": 0
}


class Gematria:
    @staticmethod
    def decode_word(word: str) -> int:
        last_idx = len(word) - 1
        result = 0
        for idx, element in enumerate(word):
            if idx == last_idx:
                result += STANDARD_ENCODING[element]
            else:
                if element in STANDARD_ENCODING:
                    result += STANDARD_ENCODING[element]
                else:
                    raise RuntimeError("Invalid letter: {}".format(element))
        return result

# user_input = input("Please enter a word or phrase you would like to find the simple english gematria of: ")
# print(f"Your word's gematria sums to: {Gematria.decode_word(user_input)}")
