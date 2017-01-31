# -*- coding: utf-8 -*-
from collections import Counter, OrderedDict
import datetime
import time
import calendar
import os
import sys
from pywinauto.application import Application
from robot.libraries.BuiltIn import BuiltIn
import pywinauto
import autoit
import getpass
from Constant import CurrencyConstant
import re

class ExtendedCustomSyExLibrary:

    ROBOT_LIBRARY_SCOPE = 'GLOBAL'

    def launch_power_express(self, version, syex_env, username):
        autoit.process_close('PowerExpress.exe')
        express_path = self.get_power_express_path(version)
        autoit.run(express_path + ' ENV:' + syex_env +' testuser:' + username + '')
        autoit.win_wait(title='Power Express')
        BuiltIn().set_test_variable('${username}', username)
        BuiltIn().set_suite_variable('${syex_env}', syex_env)

    def get_power_express_path(self, version):
        express_path = "C:\\Program Files (x86)\\Carlson Wagonlit Travel\\Power Express " + version + "\\PowerExpress.exe"
        if os.path.exists(express_path):
            return express_path
        else:
            BuiltIn().fail('Path not found')

    def get_log_path(self, test_environment, version):
        local_log_path = "C:\\Users\\"+ self.get_local_username() +"\\AppData\\Local\\Carlson Wagonlit Travel\\Power Express\\"+ version +"\\"
        citrix_log_path = "D:\\Syex_Logs\\"

        if test_environment == 'citrix':
            return citrix_log_path
        else:
            return local_log_path

    def get_local_username(self):
        return getpass.getuser()

    def select_profile(self, user_profile):
        autoit.win_wait('[REGEXPTITLE:rofil]')
        autoit.win_activate('[REGEXPTITLE:rofil]')
        select_profile_app = Application().Connect(path='PowerExpress.exe')
        select_profile_window = select_profile_app.window_(title_re=u'.*rofil')
        # select_profile_window = select_profile_app.top_window_()
        select_profile_window_list_box = select_profile_window.ListBox
        select_profile_window_list_box.Select(user_profile)
        select_profile_ok_button = select_profile_window.OK
        select_profile_ok_button.Click()
        BuiltIn().set_suite_variable('${user_profile}', user_profile)
        time.sleep(2)
        try:
            if autoit.control_command('Power Express', '[NAME:YesBtn]', 'IsVisible'):
                autoit.control_click('Power Express', '[NAME:YesBtn]')
        except Exception:
            pass

    def get_team_index_value(self, team):
        autoit.win_wait('[REGEXPTITLE:lection|Teamauswahl]')
        autoit.win_activate('[REGEXPTITLE:lection|Teamauswahl]')
        team_selection_list = self._create_power_express_app_instance().ListBox
        team_list = team_selection_list.ItemTexts()
        if team in team_list:
            return team_list.index(team)
        else:
            BuiltIn().fail("Team '" + team + "' is not found on team list")

    def set_departure_date_x_months_from_now_in_gds_format(self, number_of_months, number_of_days=0):
        future_date = self._set_future_date(number_of_months, number_of_days)
        return self._set_gds_date_format(future_date)

    def set_departure_date_x_months_from_now_in_syex_format(self, number_of_months, number_of_days=0):
        future_date = self._set_future_date(number_of_months, number_of_days)
        adjusted_date = self._adjust_weekend_to_weekday(future_date)
        return self._set_syex_date_format(adjusted_date)

    def add_days_to_current_date_in_syex_format(self, day_to_add):
        future_date = self._add_days_to_current_day(day_to_add)
        adjusted_date = self._adjust_weekend_to_weekday(future_date)
        return self._set_syex_date_format(adjusted_date)

    def set_rail_trip_date_x_months_from_now(self, number_of_months, number_of_days=0):
        future_date_formatted = self._set_future_date(number_of_months, number_of_days)
        return str(future_date_formatted.strftime('%d %B %Y'))

    def get_current_date(self):
        date_today = '{dt.month}/{dt.day}/{dt.year}'.format(dt=datetime.datetime.now())
        return str(date_today)

    def get_gds_current_date(self, remove_leading_zero='true'):
        """ 
        Returns gds current date. If you want to remove leading zero in days, set remove
        leading zero to 'true'

        | ${gds_date} = | Get Gds Current Date | remove_leading_zero=True
        """
        time_now = datetime.datetime.now().time()
        today_2pm = time_now.replace(hour=14, minute=31, second=0, microsecond=0)
        if time_now < today_2pm:
            gds_date = datetime.datetime.now() - datetime.timedelta(days=int(1))
        else:
            gds_date = datetime.datetime.now()

        if remove_leading_zero.lower() == 'true':
            return str('{dt.day}{dt:%b}'.format(dt=gds_date).upper())
        else:
            return self._set_gds_date_format(gds_date)

    def convert_date_to_gds_format(self, date, actual_date_format):
        """ 
        Example:

        | ${date} = | Convert Date To Gds Format | 11/5/2016  | %m/%d/%Y
        | ${date} = | Convert Date To Gds Format | 2016/11/09 | %Y/%m/%d

        """
        converted_date = datetime.datetime.strptime(date, actual_date_format)
        return self._set_gds_date_format(converted_date)

    def convert_date_to_syex_format(self, date, actual_date_format):
        """ 
        Example:

        | ${date} = | Convert Date To Syex Format | 11/5/2016  | %m/%d/%Y
        | ${date} = | Convert Date To Syex Format | 2016/11/09 | %Y/%m/%d

        """
        converted_date = datetime.datetime.strptime(date, actual_date_format)
        return self._set_syex_date_format(converted_date)
		
    def convert_date_to_timestamp_format(self, date, actual_date_format):
        """ 
        Example:

        | ${date} = | Convert Date To Timestamp Format | 11/5/2016  | %m/%d/%Y
        | ${date} = | Convert Date To Timestamp Format | 2016/11/09 | %Y/%m/%d

        """
        converted_date = datetime.datetime.strptime(date, actual_date_format)
        return self._set_timestamp_format(converted_date)

    def add_days_in_gds_format(self, date, day_to_add):
        converted_date = datetime.datetime.strptime(date, '%d%b')
        added_date = converted_date + datetime.timedelta(days=int(day_to_add))
        actual_date = self._adjust_weekend_to_weekday(added_date)
        return self._set_gds_date_format(actual_date)

    def add_days_in_syex_format(self, date, day_to_add):
        converted_date = datetime.datetime.strptime(date, '%m/%d/%Y')
        added_date = converted_date + datetime.timedelta(days=int(day_to_add))
        actual_date = self._adjust_weekend_to_weekday(added_date, 'add')
        return self._set_syex_date_format(actual_date)

    def subtract_days_in_gds_format(self, date, day_to_subtract):
        converted_date = datetime.datetime.strptime(date, '%d%b')
        subtracted_date = converted_date - datetime.timedelta(days=int(day_to_subtract))
        actual_date = self._adjust_weekend_to_weekday(subtracted_date)
        return str('{dt:%d}{dt:%b}'.format(dt=actual_date).upper())

    def subtract_days_in_syex_format(self, date, day_to_subtract):
        converted_date = datetime.datetime.strptime(date, '%m/%d/%Y')
        subtracted_date = converted_date - datetime.timedelta(days=int(day_to_subtract))
        actual_date = self._adjust_weekend_to_weekday(subtracted_date)
        return self._set_syex_date_format(actual_date)

    def _set_gds_date_format(self, date):
        return str('{dt:%d}{dt:%b}'.format(dt=date).upper())
		
    def _set_syex_date_format(self, date):
        return str('{dt.month}/{dt.day}/{dt.year}'.format(dt=date))
    
    def _set_timestamp_format(self, date):
        return str('{dt:%Y}-{dt:%m}-{dt:%d}'.format(dt=date))	

    def _adjust_weekend_to_weekday(self, adjusted_date, operation='subtract'):
        if str(adjusted_date.weekday()) == '5':
			if operation != 'subtract':
				return adjusted_date + datetime.timedelta(days=int(2))
			else:
				return adjusted_date - datetime.timedelta(days=int(1))
        elif str(adjusted_date.weekday()) == '6':
			if operation != 'subtract':
				return adjusted_date + datetime.timedelta(days=int(1))
			else:
				return adjusted_date - datetime.timedelta(days=int(2))
        else:
            return adjusted_date

    def _add_days_to_current_day(self, day_to_add):
        return datetime.datetime.now() + datetime.timedelta(days=int(day_to_add))

    def _set_future_date(self, number_of_months, number_of_days):
        if number_of_days > 0:
            return self._add_month_to_current_date(int(number_of_months)) + datetime.timedelta(days=int(number_of_days))
        else:
            return self._add_month_to_current_date(int(number_of_months))

    def _add_month_to_current_date(self, month_to_add):
        today = datetime.date.today()
        month = today.month - 1 + month_to_add
        year = int(today.year + month / 12)
        month = month % 12 + 1
        day = min(today.day, calendar.monthrange(year, month)[1])
        future_date = datetime.date(year, month, day)
        return future_date

    def are_there_duplicate_remarks(self, pnr_details, gds):
        # TODO
        # implement check for other gds
        """ 
        Notes:
        Amadeus - Remarks that starts with 'RM' will be checked

        """
        remarks_list = []
        pnr_log = Counter(pnr_log.strip().lower() for pnr_log in pnr_details.split('\n') if pnr_log.strip())
        for line in pnr_log:
            if pnr_log[line] > 1:
                if gds.lower() == 'amadeus' and line != 'rir *' and line.startswith("rm"):
                    print 'duplicate remarks found: ' + line
                    remarks_list.append(line)
        return True if len(remarks_list) > 0 else False

    def sort_pnr_details(self, pnr_details):
        """ 
        Example:
        ${sorted_pnr_details} = | Sort Pnr Details | {pnr_details}
        """
        return "\n".join(list(OrderedDict.fromkeys(pnr_details.split("\n"))))

    def select_value_from_dropdown_list(self, control_id, value_to_select):
        """ 
        Example:
        Select Value From Dropdown List | [NAME:ccboClass_1] | Class Code Sample
        """
        autoit.win_activate('Power Express')
        autoit.control_click('Power Express', control_id)
        autoit.send('{BACKSPACE}')
        autoit.send('{HOME}')
        while 1:
            item_combo_value = autoit.control_get_text('Power Express', control_id)
            if item_combo_value == value_to_select:
                autoit.control_focus('Power Express', control_id)
                autoit.send('{ENTER}')
                autoit.send('{TAB}')
                break
            else:
                autoit.control_focus('Power Express', control_id)
                autoit.send('{DOWN}')
            if autoit.control_get_text('Power Express', control_id) == item_combo_value:
                BuiltIn().fail("Dropdown '" + value_to_select + "' value is not found")
                break

    def get_value_from_dropdown_list(self, control_id):
        """ 
        Returns a list containing all values of the specified dropdown.
        The default dropdown value is retained.
        """
        dropdown_list = []
        autoit.win_activate('Power Express')
        autoit.control_focus('Power Express', control_id)
        original_value = autoit.control_get_text('Power Express', control_id)
        autoit.send('{BACKSPACE}')
        autoit.send('{HOME}')
        while 1:
            item_combo_value = autoit.control_get_text('Power Express', control_id)
            dropdown_list.append(item_combo_value)
            autoit.send('{DOWN}')
            if autoit.control_get_text('Power Express', control_id) == item_combo_value:
                break
        autoit.send('{BACKSPACE}')
        autoit.send('{HOME}')
        for item in dropdown_list:
            if original_value == autoit.control_get_text('Power Express', control_id):
                break
            else:
                autoit.send('{DOWN}')
        return dropdown_list

    def _create_power_express_app_instance(self):
        app = Application().Connect(path='PowerExpress.exe')
        power_express_app_instance = app[u'WindowsForms10.Window.8.app.0.3309ded_r17_ad1']
        power_express_app_instance.Wait('ready')
        return power_express_app_instance

    def select_tab_control(self, tab_control_value):
        """ 
        Example:
        Select Tab Control | Fare 1
        """
        tab_control = self._create_power_express_app_instance().TabControl
        tab_control.Select(tab_control_value)

    def select_value_from_combobox(self, combobox_value):
        combobox = self._create_power_express_app_instance().ComboBox
        if self._is_control_enabled(combobox):
            combobox.Select(u''.join(combobox_value))

    def select_tst_fare(self, fare_number):
        # For Rail Panel
        listbox = self._create_power_express_app_instance().ListBox
        listbox.Select(u'  %s' % fare_number)

    def select_value_from_listbox(self, listbox_value):
        listbox = self._create_power_express_app_instance().ListBox
        listbox.Select(u''.join(listbox_value))

    def get_value_from_combobox(self, control):
        autoit.control_click('Power Express', control)
        combobox = self._create_power_express_app_instance().ComboBox
        return combobox.ItemTexts()

    def _is_control_enabled(self, control):
        return True if control.IsEnabled() else False

    def is_tab_visible(self, tab_value):
        tab_control_texts = self._create_power_express_app_instance().TabControl.Texts()
        if tab_value in tab_control_texts:
            return True
        else:
            return False

    def get_currency(self, country_code):
        return getattr(CurrencyConstant, country_code.lower())

    def get_percentage_value(self, number):
        percentage = (float(number) / float(100))
        return percentage

    def convert_to_float(self, value, precision=2):
        """ 
        Converts value to float. By default precision used is 2.

        Example:
        ${convered_value} = | Convert To Float | 15.0909 | precision=5

        """
        try:
            return "{:.{prec}f}".format(float(value), prec=precision)
        except Exception as error:
            raise ValueError(repr(error))

    def get_visible_tab(self):
        """ 
        Returns tabs in list data type

        """
        tab_control = self._create_power_express_app_instance().TabControl
        tab_list = []
        tab = tab_control.WrapperObject()
        for tab_item in range(0, tab.TabCount()):
            tab_list.append(tab.GetTabText(tab_item).strip())
        return tab_list

    def get_string_matching_regexp(self, reg_exp, data):
        m = re.search(reg_exp, data)
        return m.group(0)

    def get_substring_using_marker(self, whole_string, first_marker, end_marker=None):
        """ 
        Returns substring or between string from first marker and end marker

        """
        start = whole_string.find(first_marker) + len(first_marker)
        end = whole_string.find(end_marker, start)
        return whole_string[start:end].strip()

    def get_minimum_value_from_list(self, list_):
        """ 
        Returns least / minimum value from list

        """
        return min(list_)

    def get_required_flight_details(self, raw_flight_details, city_pair_marker):
        raw_flight_details = raw_flight_details.strip()
        flight_details_line = raw_flight_details.find(city_pair_marker)
        fare_line_length = flight_details_line + len(city_pair_marker)
        raw_flight_details = raw_flight_details[:fare_line_length]
        flight_details_list = raw_flight_details.split(" ")
        flight_details_list = [x for x in flight_details_list if x]
        del flight_details_list[0]
        return flight_details_list
