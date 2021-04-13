[![tagged-release](https://github.com/sharkoz/keypirinha-outlook_calendar/actions/workflows/main.yml/badge.svg)](https://github.com/sharkoz/keypirinha-outlook_calendar/actions/workflows/main.yml)

# Keypirinha Plugin: outlook_calendar

This is outlook_calendar, a plugin for the
[Keypirinha](http://keypirinha.com) launcher.

It allows you to view upcoming events in your outlook calendar,
and to join a teams meeting if available 

Are you tired of always looking for the link of your next teams meeting in outlook ?
With this plugin, you can see from keypirinha your next outlook appointments for the day,
and join your MS Teams meetings in just one keypress.


## Recommended install

Install directly from Keypirinha using ueffel's [PackageControl](https://github.com/ueffel/Keypirinha-PackageControl)


## Manual Install

You can download the latest `.keypirinha-package` in the `releases` section :
https://github.com/sharkoz/keypirinha-outlook_calendar/releases

Once the `outlook_calendar.keypirinha-package` file is installed,
move it to the `InstalledPackage` folder located at:

* `Keypirinha\portable\Profile\InstalledPackages` in **Portable mode**
* **Or** `%APPDATA%\Keypirinha\InstalledPackages` in **Installed mode** (the
  final path would look like
  `C:\Users\%USERNAME%\AppData\Roaming\Keypirinha\InstalledPackages`)


## Usage

Open Keypirinha and select the `Calendar` entry.

View your upcoming events, and if they have a linked teams meeting, join the meeting
by highlighting the event with the up/down arrows and pressing `Enter`


## Configuration

You can configure these settings :

* label : The catalog entry name (if you want to type `ocal` instead of `Calendar` to invoke the plugin for example)
* max_results : To limit the number of events returned
* max_days : To limit the number of days in the future scanned for events


## Change Log


### v1.1
* Fix missing events
* Added configuration options

### v1.0

* First version


## License

This package is distributed under the terms of the MIT license.


## Contribute

1. Check for open issues or open a fresh issue to start a discussion around a
   feature idea or a bug.
2. Fork this repository on GitHub to start making your changes to the **dev**
   branch.
3. Send a pull request.
4. Add yourself to the *Contributors* section below (or create it if needed)!
