# %%
from ics import Calendar, Event
from imap_tools import MailBox, AND, A
import typer
import datetime as dt
from bs4 import BeautifulSoup
import re
import logging
import sys
import os
logger = logging.getLogger("rail-ics")
app = typer.Typer()
# %%


def _fetch(output: str, show_past: bool, imap_server: str, imap_user: str, imap_pass: str, imap_folder: str):

    logger.debug("start fetch new calender")
    c = Calendar()
    # print("BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:rail-ics.py")
    with MailBox(imap_server).login(imap_user, imap_pass, initial_folder=imap_folder) as mailbox:
        for msg in mailbox.fetch(A(from_="12306@rails.com.cn", date_gte=dt.date(2020, 12, 1)), bulk=True, reverse=True):
            logger.debug(f"{msg.subject} {msg.date}")
            soup = BeautifulSoup(msg.html, "html.parser")
            try:
                ticket = soup.select(
                    "body>table>tr>td>table>tr>td>div>div")[0].text.strip()
                g = re.match(
                    "^\d\.\w*?，(\d{4})年(\d{2})月(\d{2})日(\d{2}):(\d{2})开，.*$", ticket)
                y, m, d, H, M = [int(i) for i in g.groups()]
            except Exception as e:
                logger.warning(f"parse email error,{msg}")
                logger.exception(e)
                continue
            min_query_date = dt.datetime.now(
                tz=msg.date.tzinfo)-dt.timedelta(days=61)
            if min_query_date > msg.date:
                break
            begin = dt.datetime(y, m, d, H, M, 0)
            if show_past == False:
                min_date = dt.datetime.now()-dt.timedelta(days=1)
                if begin < min_date:
                    continue

            end = begin+dt.timedelta(hours=1)
            e = Event()
            e.name = ticket
            e.begin = begin
            e.end = end
            logger.debug(f"New Event: {e}")
            c.events.add(e)
    # print("END:VCALENDAR")
    if output is None:
        folder = "ics"
        output = os.path.join(folder, f"{imap_user}.ics")
        os.makedirs(folder, exist_ok=True)
    if output:
        with open(output, "w") as fp:
            fp.write(str(c))


@app.command()
def fetch(output: str = None, show_past: bool = False, imap_server: str = None, imap_user: str = None,
          imap_pass: str = None, imap_folder: str = "INBOX"):
    _fetch(output=output, show_past=show_past, imap_server=imap_server,
           imap_user=imap_user, imap_pass=imap_pass, imap_folder=imap_folder)


@app.command()
def server():
    pass


if __name__ == "__main__":
    app()
