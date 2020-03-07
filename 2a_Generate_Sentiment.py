from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
import xlrd
import xlwt

data_path = "D:\MCM2020\pacifier.xlsx"

analyser = SentimentIntensityAnalyzer()
f = xlwt.Workbook()
wb = xlrd.open_workbook(filename=data_path)
sheet1 = wb.sheet_by_index(0)
sheet2 = f.add_sheet('page1', cell_overwrite_ok=True)
col_r = sheet1.col_values(12)
col_r2 = sheet1.col_values(13)

for i in range(0, len(col_r)):
    sheet2.write(i, 0, analyser.polarity_scores(str(col_r[i]) + str(col_r2[i]))['compound'])
    if not (i % 100):
        print(i)
# score = analyser.polarity_scores("Disappointment with dryer. I purchased it because it was supposed to be quiet. It's every bit as loud as my old dryer. It's heavy, cumbersome, hard to manage. I kept turning it off because of the location of the buttons on the handle (I didn't have that problem with my old dryer). It kept sucking my hair in the motor area. <br />BUT, I do think there's something to this ion thing. My hair seemed softer and straighter - no frizzies. It also seemed to dry faster. So, I am now on a quest to find a ion dryer that is light, quiet and easy to manage - oh, and doesn't eat my hair.")
# print(score['compound'])
f.save('D:\\MCM2020\\pacifier_sentiment.xls')
