import re

## Функция отчистки текста от всякой всячины
def optimazeText(text):
    optimazed_text = re.sub(r"\n", " ", text)
    optimazed_text = re.sub(
        "[\U00000000-\U0000002B|\U0000003A-\U0000040F|\U00000450-\U0010FFFF|\U0000002F]", " ", optimazed_text)
    optimazed_text = re.sub(r"\t", "", optimazed_text)
    optimazed_text = re.sub("\s+", " ", optimazed_text)
    optimazed_text = re.sub("-", " ", optimazed_text)

    return optimazed_text.replace("NaN", "").replace("NaT", "").replace("Unnamed", "").replace("None", "").replace("\\", "")
