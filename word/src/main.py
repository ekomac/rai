from config import CONFIG
import docx


def main():
    print(CONFIG)
    doc = docx.Document('test.docx')
    doc.paragraphs.append('Hello World')


if __name__ == '__main__':
    main()
