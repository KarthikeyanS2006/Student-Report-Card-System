import discord
import os

# Replace with your actual bot token from the Discord Developer Portal
TOKEN = "MTQ1MTQxMzI2NzQ3NzAzNzEzOQ.GoFRHn.8jkDS_65rhmn7Ng5srYfxYRoSE-65R5NcfhMFM"

# The application ID is not used in the code, but you can keep it for reference
APPLICATION_ID = "1451413267477037139"

# Only enable necessary intents
intents = discord.Intents.default()
intents.message_content = True  # Enable message content intent

client = discord.Client(intents=intents)

# Import your automated ReportCardSystem
from app import ReportCardSystem

system = ReportCardSystem()

@client.event
async def on_ready():
    print(f'{client.user} has connected to Discord!')

@client.event
async def on_message(message):
    if message.author == client.user:
        return

    if message.content.startswith('!addstudent'):
        # Extract user input
        # Format: !addstudent 101 John 90 85 80 75 70
        parts = message.content.split()
        if len(parts) != 8:
            await message.channel.send("Please use the correct format: !addstudent <ID> <Name> <Math> <Science> <English> <Social> <Computer>")
            return

        student_id = parts[1]
        student_name = parts[2]
        math = float(parts[3])
        science = float(parts[4])
        english = float(parts[5])
        social = float(parts[6])
        computer = float(parts[7])

        # Call the automated add_student function
        result = system.add_student_automated(student_id, student_name, math, science, english, social, computer)
        await message.channel.send(result)

    elif message.content.startswith('!viewstudents'):
        # View all students
        result = system.view_all_students_automated()
        await message.channel.send(f"``````")

    elif message.content.startswith('!searchstudent'):
        # Search student by ID
        parts = message.content.split()
        if len(parts) != 2:
            await message.channel.send("Please use the correct format: !searchstudent <ID>")
            return
        student_id = parts[1]
        result = system.search_student_automated(student_id)
        await message.channel.send(result)

    elif message.content.startswith('!generatereport'):
        # Generate report card for a student
        parts = message.content.split()
        if len(parts) != 2:
            await message.channel.send("Please use the correct format: !generatereport <ID>")
            return
        student_id = parts[1]
        result = system.generate_report_card_automated(student_id)
        await message.channel.send(result)

client.run(TOKEN)
