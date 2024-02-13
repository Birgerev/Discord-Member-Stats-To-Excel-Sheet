# This example requires the 'message_content' intent.

import discord
import datetime
from monthdelta import monthdelta
import xlsxwriter

token = ''
channelId = 694254855761428494

workbook = xlsxwriter.Workbook('server_statistics.xlsx')
worksheet = workbook.add_worksheet()

interval = monthdelta(1)

intents = discord.Intents.default()
intents.message_content = True
client = discord.Client(intents=intents)

@client.event
async def on_ready():
    print(f'We have logged in as {client.user}')
    channel = client.get_channel(channelId)

    date = channel.created_at
    lastDate = datetime.datetime.now(datetime.timezone.utc)
    row = 1
    column = 1
    user_columns = {}

    while (date <= lastDate):
        print(date.strftime("%Y-%m"))
        intervalEnd = date + interval
        user_message_amount = {"jonas" : 5}

        #async for message in channel.history(limit=50000, after=date, before=(intervalEnd)):
        #    sender = message.author.name
        #    if sender in user_message_amount.keys():
        #        user_message_amount[sender] += 1
        #    else:
        #        user_message_amount[sender] = 1

        worksheet.write(row, 0, date.strftime("%Y-%m"))
        for user in user_message_amount:
            print(f'{user} : {user_message_amount[user]}')
            userColumn = column
            if sender in user_columns.keys():
                userColumn = user_columns[sender]
            else:
                user_columns[sender] = userColumn
                column +=1
            worksheet.write(row, userColumn, date.strftime("%Y-%m"))

        date += interval
        row += 1
    workbook.close()



client.run(token)