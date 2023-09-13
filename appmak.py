#!/usr/bin/env python3

import ctypes
import win32com.client
import argparse
import re
from datetime import datetime, timedelta


def parse_time(s: str) -> datetime:
    # this is the verbose datetime as we know it
    today = datetime.now()
    year = f"{today.year:04d}"
    month = f"{today.month:02d}"
    day = f"{today.day:02d}"
    pattern_full = re.compile(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}")
    pattern_time = re.compile(r"\d{2}:\d{2}:\d{2}")
    pattern_time_short = re.compile(r"\d{4}")
    if (re.match(pattern_full, s)):
        return datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
    if (re.match(pattern_time, s)):
        return datetime.strptime(f"{year}-{month}-{day} " + s, "%Y-%m-%d %H:%M:%S")
    if (re.match(pattern_time_short, s)):
        return datetime.strptime(f"{year}-{month}-{day} " + s[:2] + ":" + s[2:] + ":00", "%Y-%m-%d %H:%M:%S")


def parse_args():
    """parse the commandline arguments that are important for the app"""
    parser = argparse.ArgumentParser(
        description="Create an Outlook appointment.")

    # Required arguments
    parser.add_argument("start_time", type=lambda s: parse_time(
        s), help="Start time in the format 'YYYY-MM-DD HH:MM:SS' or today short hand with 'HH:MM:SS' or short hand for only the hour and minute 'HHMM'")

    # Optional arguments
    # Attendees are a feature for the future
    # parser.add_argument("-a", "--attendees", nargs='+', default=[], help="List of attendees (space-separated)")
    parser.add_argument("-b", "--body", type=str, default="",
                        help="Body text of the appointment")
    parser.add_argument("-d", "--display", default=False,
                        action="store_true", help="Whether to display the appointment after creation")
    parser.add_argument("-v", "--verbose", default=False, action="store_true",
                        help="Whether to display verbose output in the command line")
    parser.add_argument("-e", "--end_time", type=lambda s: parse_time(s),
                        default=None, help="End time in the format 'YYYY-MM-DD HH:MM:SS'")
    parser.add_argument("-t", "--title", type=str,
                        default="Placeholder", help="Title of the appointment")
    parser.add_argument("-l", "--label", type=str, default=None,
                        help="Label for title: Default NONE, Available: t [TASK] o [ORGA] g [GYM], Any: ['Label']")

    args = parser.parse_args()

    # if there is no end_time we add 30 minutes and that is it
    if (args.end_time == None):
        delta = timedelta(minutes=30)
        args.end_time = args.start_time + delta

    # short hands for common labels
    if (args.label != None):
        if (len(args.label) == 1):
            if (args.label == "t"):
                args.title = f"[TASK]: {args.title}"
            elif (args.label == "o"):
                args.title = f"[ORGA]: {args.title}"
            elif (args.label == "g"):
                args.title = f"[GYM]: {args.title}"
            else:
                print("INFO: Label not registered")
                args.title = f"[{args.label}]: {args.title}"
        else:
            args.title = f"[{args.label}]: {args.title}"

    # print(args.start_time, args.end_time)
    return args


def make_appointment(start_time, end_time, title, display, body, attendees=[]):
    """modify the otf template so that the arguments are written in the right spott
    Args:
        start_time: date_time
        end_time: date_time
        title: string
        body: string
        attendies: list of strings
    Return:
        otf file with the right strukture
    """
    # Create an instance of the Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application")

    # Create an appointment item
    appointment = outlook.CreateItem(1)  # 1 corresponds to olAppointmentItem

    # Set appointment properties
    appointment.Subject = title
    appointment.Body = body
    appointment.Start = start_time
    appointment.End = end_time
    # this is a thing for the future
    # appointment.RequiredAttendees = attendees

    # Save the appointment
    appointment.Save()

    if (display):
        appointment.Display()


def verbose(start_time, end_time, title, display, body):
    print("===Verbose Ouput===")
    print("Start:", start_time)
    print("End:", end_time)
    print("Title:", title)
    print("Body:", body)
    print("Display:", display)
    print("===DONE===")


def main():
    args = parse_args()
    make_appointment(args.start_time, args.end_time,
                     args.title, args.display, args.body)
    if (args.verbose):
        verbose(args.start_time, args.end_time,
                args.title, args.display, args.body)


if __name__ == "__main__":
    main()
