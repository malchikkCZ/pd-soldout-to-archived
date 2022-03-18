import re
import unicodedata


class Handleizer:

    def __init__(self):
        pass

    @classmethod
    def run(cls, string):
        data = [char for char in unicodedata.normalize("NFKD", string) if unicodedata.category(char) != "Mn"]
        ascii_string = "".join(data)

        handleized_title = re.sub(
            r"^-", "", re.sub(
                r"-$", "", re.sub(
                    r"[^a-z0-9]+", "-", ascii_string.lower()
                )
            )
        )

        return handleized_title


if __name__ == '__main__':
    pass
