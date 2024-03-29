from datetime import datetime
import datetime as dt
from urllib.request import unquote

import keypirinha as kp
import keypirinha_util as kpu
import comtypes.client as com
import re, os


class Outlook_cal(kp.Plugin):
    """
    Outlook calendar

    View upcoming events in your outlook calendar, and join teams meeting
    if available 

    """

    ITEMCAT = kp.ItemCategory.USER_BASE + 1

    def __init__(self):
        super().__init__()
        self.info('init')

    def on_start(self):
        self._read_config()

    def on_catalog(self):
        self.set_catalog([
            self.create_item(
                category=kp.ItemCategory.KEYWORD,
                label=self.settings.get("label", "main", "Calendar", True),
                short_desc="View upcoming Outlook meetings",
                target="Calendar",
                args_hint=kp.ItemArgsHint.REQUIRED,
                hit_hint=kp.ItemHitHint.IGNORE
            )
        ])

    def on_suggest(self, user_input, items_chain):
        if not self.settings.get_bool("always_suggest", "main", False) and (not items_chain or items_chain[0].category() != kp.ItemCategory.KEYWORD):
            return
        date = datetime.now()
        cal = self.__get_calendar(date,date+dt.timedelta(days=int(self.settings.get("max_days", "main", 5, True))))
        suggestions = self.__compose_suggestions(cal, user_input)
        self.set_suggestions(suggestions, kp.Match.ANY, kp.Sort.LABEL_ASC )

    def on_execute(self, item, action):
        kpu.shell_execute("msteams"+item.data_bag())

    def on_activated(self):
        pass

    def on_deactivated(self):
        pass

    def on_events(self, flags):
        self.on_catalog()

    def _read_config(self):
        self.settings = self.load_settings() 
        self.kpsettings = kp.settings()

    def __compose_suggestions(self, cal, user_input) -> []:
        # Label of outlook response status
        status = {1:"Organizer", 2: "Tentative", 3: "Accepted", 4:"Declined", 5: "Pending"}
        # Icon of outlook response status
        icons={1:"organizer", 2:"maybe", 3:"ok", 4:"ko", 5:"maybe"}
        suggestions = []
        nb=0
        for app in cal:
            # If no user input, display everything. Otherwise, lookup in the subject field
            # Added the "space" condition to allow for quick display when always_suggest is true :
            # In that last case, opening the window and only pressing "space" will display your next appointments
            if len(user_input) < 1 or app.subject.lower().find(user_input.lower())>=0 or user_input==" ":
                nb = nb+1
                link="None"
                desc= app.location
                body = str(app.body)
                # Check for MS teams link
                srch = re.search(r":\/\/teams.[\w\d:#@%\/;$()~_?\+-=\\\.&]*",body)
                # If not found, check a outlook safelink protected link
                if not srch:
                    for link in re.findall(r"safelinks\.protection\.outlook\.com[\S]*\?url=([^&]*)",body):
                        srch = re.search(r":\/\/teams.[\w\d:#@%\/;$()~_?\+-=\\\.&]*",unquote(link))
                        if srch:
                            break
                if srch:
                    link = srch.group()
                    desc += " - Press [Enter] to open in Teams"
                else:
                    self.dbg('No teams link found.')
                new = self.__create_suggestion_item(
                    nb,
                    icons.get(app.responseStatus, "maybe"),
                    str(app.start)[:-3] + "-" + str(app.end)[-8:-3] + "  " + app.subject + " - " + status.get(app.responseStatus, ""),
                    desc,
                    link)
                suggestions.append(new)
                if nb>=int(self.settings.get("max_results", "main", self.kpsettings.get("max_height", "gui", 10, True), True)):
                    break
        return suggestions  

    def __create_suggestion_item(self, id, icon, label: str, short_desc: str, link: str):
        if self.should_terminate():
            return
        return self.create_item(
            category=self.ITEMCAT,
            label=label,
            short_desc=short_desc,
            target=str(id),
            args_hint=kp.ItemArgsHint.FORBIDDEN,
            hit_hint=kp.ItemHitHint.IGNORE,
            icon_handle=self.load_icon("res://"+self.package_full_name()+"/"+icon+".ico"),
            data_bag=link
        )

    def __get_calendar(self,begin,end):
        outlook = com.CreateObject("Outlook.Application", dynamic=True).GetNamespace('MAPI')
        calendar = outlook.getDefaultFolder(9).Items
        calendar.IncludeRecurrences = True
        calendar.Sort('[Start]')
        date_format = self.settings.get("date_format", "main", '%d-%m-%Y', True)
        restriction = "[END] >= '" + begin.strftime(date_format+' %H:%M') + "' AND [END] <= '" + end.strftime(date_format) + "'"
        calendar = calendar.Restrict(restriction)
        return calendar
