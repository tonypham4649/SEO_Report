from abc import ABC, abstractmethod


class LanguageText(ABC):
    """abstract class
    """
    @abstractmethod
    def translate_to_lang(self, intext: str):
        pass


class EnglishText(LanguageText):

    def translate_to_lang(self, intext: str):
        return intext


class JapaneseText(LanguageText):
    """just a explain
    """

    def __init__(self) -> None:
        self._langDict = {
            'hi': 'お早うございます',
            'bye': 'さよなら'
        }

    def translate_to_lang(self, intext: str):
        return self._langDict.get(intext)


class VietnameseText(LanguageText):
    def __init__(self) -> None:
        self._langDict = {
            'hi': 'chào',
            'bye': 'chào'
        }

    def translate_to_lang(self, intext: str):
        return self._langDict.get(intext)


def client_code(language='English') -> LanguageText:
    """main factory that take the input and then return the class as output
    """
    mapping = {
        'English': EnglishText,
        'Japanese': JapaneseText,
        'Vietnamese': VietnameseText
    }

    return mapping[language]()


if __name__ == '__main__':
    v = client_code('Vietnamese')

    print(v.translate_to_lang('hi'))
    print(v.translate_to_lang('bye'))

    v = client_code('Japanese')

    print(v.translate_to_lang('hi'))
    print(v.translate_to_lang('bye'))
