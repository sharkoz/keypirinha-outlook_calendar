from datetime import datetime
import datetime as dt

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
        pass

    def on_catalog(self):
        self.set_catalog([
            self.create_item(
                category=kp.ItemCategory.KEYWORD,
                label="Calendar",
                short_desc="View upcoming Outlook meetings",
                target="Calendar",
                args_hint=kp.ItemArgsHint.REQUIRED,
                hit_hint=kp.ItemHitHint.IGNORE
            )
        ])

    def on_suggest(self, user_input, items_chain):
        if not items_chain or items_chain[0].category() != kp.ItemCategory.KEYWORD:
            return
        date = datetime.now()
        cal = self.__get_calendar(date,date+dt.timedelta(days=5))
        suggestions = self.__compose_suggestions(cal)
        self.set_suggestions(suggestions, kp.Match.ANY, kp.Sort.LABEL_ASC )

    def on_execute(self, item, action):
        kpu.shell_execute("msteams"+item.target())

    def on_activated(self):
        pass

    def on_deactivated(self):
        pass

    def on_events(self, flags):
        pass

    def __compose_suggestions(self, cal) -> []:
        status = {1:"Organizer", 2: "Tentative", 3: "Accepted", 4:"Declined", 5: "Pending"}
        suggestions = []
        nb=0
        for app in cal:
            nb = nb+1
            link="None"
            desc= app.location
            srch = re.search(r"://teams.[\w\d:#@%/;$()~_?\+-=\\\.&]*",str(app.body))
            if srch:
                link = srch.group()
                desc += " - Press [Enter] to open in Teams"
            new = self.__create_suggestion_item(str(app.start)[:-3] + "-" + str(app.end)[-8:-3] + "  " + app.subject + " - " + status.get(app.responseStatus, ""), desc, link)
            suggestions.append(new)
            if nb>10:
                break
        return suggestions  

    def __create_suggestion_item(self, label: str, short_desc: str, target: str):
        return self.create_item(
            category=self.ITEMCAT,
            label=label,
            short_desc=short_desc,
            target=target,
            args_hint=kp.ItemArgsHint.FORBIDDEN,
            hit_hint=kp.ItemHitHint.IGNORE,
        )

    def __get_calendar(self,begin,end):
        outlook = com.CreateObject("Outlook.Application", dynamic=True).GetNamespace('MAPI')
        calendar = outlook.getDefaultFolder(9).Items
        calendar.IncludeRecurrences = True
        calendar.Sort('[Start]')
        restriction = "[END] >= '" + begin.strftime('%d/%m/%Y %H:%M') + "' AND [END] <= '" + end.strftime('%d/%m/%Y') + "'"
        calendar = calendar.Restrict(restriction)
        return calendar