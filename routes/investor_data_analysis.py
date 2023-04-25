from datetime import datetime

import pandas as pd

from fastapi import APIRouter, Request
from fastapi.responses import FileResponse
from portal_statement_files.portal_statement_create import create_pdf
from config.db import db
from bson import ObjectId

data_analysis = APIRouter()

# MONGO COLLECTIONS


investors = db.investors


@data_analysis.get("/get_all_investors_to_analyse")
def get_all_investors_to_analyse():
    investor_list = list(investors.find())
    for investor in investor_list:
        investor['_id'] = str(investor['_id'])
    # print("investor_list", investor_list)
    investor_investments = []
    for investor in investor_list:
        # if investor['investments'] and investor['investments'] != []:
        if investor['trust']:
            for investment in investor['trust']:
                investment['investor_acc_number'] = investor['investor_acc_number']
                investment['investor_name'] = investor['investor_name']
                investment['investor_surname'] = investor['investor_surname']
                investment['investor_id_number'] = investor['investor_id_number']
                investment['opportunity_code'] = investment['opportunity_code']
                investment['Category'] = investment['Category']
                # convert investment_amount to float
                investment['investment_amount'] = float(investment['investment_amount'])
                # convert deposit_date and release_date to datetime if they are not equal to "", if the are equal to
                # "" then set them to None
                if investment['deposit_date'] != "":
                    # replace "/" with "-" in deposit_date
                    investment['deposit_date'] = investment['deposit_date'].replace("/", "-")
                    investment['deposit_date'] = datetime.strptime(investment['deposit_date'], "%Y-%m-%d")
                else:
                    investment['deposit_date'] = None
                if investment['release_date'] != "":
                    # replace "/" with "-" in release_date
                    investment['release_date'] = investment['release_date'].replace("/", "-")
                    investment['release_date'] = datetime.strptime(investment['release_date'], "%Y-%m-%d")
                else:
                    investment['release_date'] = None
                investor_investments.append(investment)

    # filter out investor_investments that have a Category equals to "Southwark"
    investor_investments = [x for x in investor_investments if x['Category'] != "Southwark"]

    final_investments = []

    # loop investor_investments and append only "investment_amount", "deposit_date", "release_date",
    # "investor_acc_number", "investor_name", "investor_surname", "investor_id_number" to final_investments
    for investment in investor_investments:
        if len(investment['investor_id_number']) == 13:
            investment['investor_id_number'] = investment['investor_id_number']
            year = investment['investor_id_number'][:2]
            month = investment['investor_id_number'][2:4]
            day = investment['investor_id_number'][4:6]
            if int(year) <= 23:
                year = "20" + year
            else:
                year = "19" + year
            # print("year", year) take year, month and day and convert them to datetime and then calculate the age of
            # the investor in years only and save it to age variable
            age = int((datetime.now() - datetime(int(year), int(month), int(day))).days / 365)
            # print("age", age)
            # set investor_age to age
            investment['investor_age'] = age

            # create a variable called sex, if investment['investor_id_number'][6:10] < "5000" as an integer then sex
            # equals female else sex equals male

            if int(investment['investor_id_number'][6:10]) < 5000:
                investment['investor_sex'] = 'female'
            else:
                investment['investor_sex'] = 'male'
        else:
            investment['investor_id_number'] = None
            investment['investor_age'] = None
            investment['investor_sex'] = "Legal Entity"

        final_investments.append({
            "investment_amount": investment['investment_amount'],
            "deposit_date": investment['deposit_date'],
            "release_date": investment['release_date'],
            "investor_acc_number": investment['investor_acc_number'],
            "investor_name": investment['investor_name'],
            "investor_surname": investment['investor_surname'],
            "investor_id_number": investment['investor_id_number'],
            "investor_age": investment['investor_age'],
            "investor_sex": investment['investor_sex'],
            "opportunity_code": investment['opportunity_code'],
            "Category": investment['Category']
        })

    # save final_investments to a csv file
    df = pd.DataFrame(final_investments)
    # print("df", df)
    # convert deposit_date to deposit_month and deposit_year to string as two columns
    df['deposit_month'] = df['deposit_date'].dt.strftime('%m')
    df['deposit_year'] = df['deposit_date'].dt.strftime('%Y')
    # convert release_date to release_month and release_year to string as two columns
    df['release_month'] = df['release_date'].dt.strftime('%m')
    df['release_year'] = df['release_date'].dt.strftime('%Y')

    # group by deposit_month and deposit_year and sum the investment_amount
    df2 = df.groupby(['deposit_year', 'deposit_month'])['investment_amount'].sum().reset_index()
    # for each year, have a cumulative total of the investment_amount
    df2['cumulative_total'] = df2.groupby('deposit_year')['investment_amount'].cumsum()

    # convert both the sum and the cumulative total to integers
    df2['investment_amount'] = df2['investment_amount'].astype(int)
    df2['cumulative_total'] = df2['cumulative_total'].astype(int)

    # create a list of months by name
    months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
              'November', 'December']

    # for each row in df2, get the month number and insert with the month name called month_name
    df2['month_name'] = df2['deposit_month'].apply(lambda x: months[int(x) - 1])

    # convert the month and year to integers
    df2['deposit_month'] = df2['deposit_month'].astype(int)
    df2['deposit_year'] = df2['deposit_year'].astype(int)

    # from df get the average age of investors, ignore the null values, if the result is NaN then set it to the average
    df3 = df.groupby(['deposit_year', 'deposit_month'])['investor_age'].mean().reset_index()
    # in df3 replace the NaN values with the average of investor_age
    df3['investor_age'] = df3['investor_age'].fillna(df3['investor_age'].mean())
    print("df3", df3)
    # add the investor age to df2 as an integer
    df2['investor_age'] = df3['investor_age'].astype(int)

    # df.to_csv("final_investments.csv", index=False)

    return df2.to_dict('records')
    # return final_investments
