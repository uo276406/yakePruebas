import yake
import xlsxwriter
from spacy.lang.en.stop_words import STOP_WORDS

data = [['USHist_', 32], ['chapter', 36], ['Cambridge_IGCSE_History_', 10], ['From yesterday to tomorrow _ history and citizenship education_glossary_', 6]]


def write_row_excel(worksheet, chapter, tp, fn, fp, precision, recall, f1score, row_num):
    worksheet.write(row_num+1, 0, chapter)
    worksheet.write(row_num+1, 1, tp)
    worksheet.write(row_num+1, 2, fp)
    worksheet.write(row_num+1, 3, fn)
    worksheet.write(row_num+1, 4, precision)
    worksheet.write(row_num+1, 5, recall)
    worksheet.write(row_num+1, 6, f1score)


def main():
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('./resultsYakeCustom.xlsx')

    for name, iterations in data:

        if name == 'From yesterday to tomorrow _ history and citizenship education_glossary_':
            worksheet = workbook.add_worksheet(name='From yesterday to tomorrow')
        else:
            worksheet = workbook.add_worksheet(name=name)

        worksheet.write(0, 0, "text name")
        worksheet.write(0, 1, "tp")
        worksheet.write(0, 2, "fp")
        worksheet.write(0, 3, "fn")
        worksheet.write(0, 4, "precision")
        worksheet.write(0, 5, "recall")
        worksheet.write(0, 6, "f1score")

        for i in range(iterations):
            if name == 'chapter':
                i = i+1

            name = "From yesterday to tomorrow _ history and citizenship education_glossary_"
            i = 0

            f = open("./data/History/dataSet/" + name + str(i) + ".txt", 'r', encoding='utf8')
            text = f.read()
            f.close()

            language = "en"
            max_ngram_size = 3
            deduplication_threshold = 0.9
            deduplication_algo = 'seqm'
            windowSize = 3
            numOfKeywords = 500

            custom_kw_extractor = yake.KeywordExtractor(lan=language, n=max_ngram_size,
                                                        dedupLim=deduplication_threshold, dedupFunc=deduplication_algo,
                                                        windowsSize=windowSize, top=numOfKeywords, features=None)
            keywords = custom_kw_extractor.extract_keywords(text)

            keywords_found = []
            for k, s in keywords:
                keywords_found.append(k)
            print(keywords_found)
            key = open("./data/History/dataSet/" + name + str(i) + ".key", 'r', encoding='utf8')
            keywords_real = key.readlines()
            key.close()

            tp = 0
            for k in keywords_real:
                keyword = k.lower().strip()
                keyword_processed = keyword.split()
                keyword_to_check = ""
                for elem in keyword_processed:
                    if elem not in STOP_WORDS:
                        keyword_to_check += (elem + " ")
                if keyword_to_check.strip() in keywords_found:
                    tp += 1

            fn = len(keywords_real)-tp
            fp = len(keywords_found)-tp
            precision = tp / (tp + fp)
            recall = tp / (tp + fn)
            f1score = (2 * tp) / (2 * tp + fp + fn)

            print("Keywords found for " + name + str(i))
            print(keywords_found)
            print("TP:" + str(tp))
            print("FN:" + str(fn))
            print("FP:" + str(fp))
            print("Precision: " + str(precision))
            print("Recall: " + str(recall))
            print("F1Score: " + str(f1score))
            print()
            write_row_excel(worksheet, name + str(i), tp, fn, fp, precision, recall, f1score, i)

    workbook.close()


if __name__ == "__main__":
    main()
