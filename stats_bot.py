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
    emptyColumn = 1
    user_columns = {}

    print("Starting message logging, this will take a while...")
    while (date <= lastDate):
        print(date.strftime("%Y-%m"))
        intervalEnd = date + interval
        user_message_amount = {}

        async for message in channel.history(limit=50000, after=date, before=(intervalEnd)):
            sender = message.author.name
            if sender in user_message_amount.keys():
                user_message_amount[sender] += 1
            else:
                user_message_amount[sender] = 1

        worksheet.write(row, 0, date.strftime("%Y-%m"))
        for user in user_message_amount:
            print(f'{user} : {user_message_amount[user]}')

            userColumn = emptyColumn
            if user in user_columns.keys():
                #Use user column
                userColumn = user_columns[user]
            else:
                #create user column
                user_columns[user] = userColumn
                worksheet.write(0, userColumn, user)
                emptyColumn +=1

            #Fill individual cell with message amount
            worksheet.write(row, userColumn, user_message_amount[user])

        date += interval
        row += 1
    workbook.close()



client.run(token)