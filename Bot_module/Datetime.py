import datetime

class DateTimeUtility:
    """
    A utility class to perform various date and time operations.
    """

    def get_current_date(self, format_string="%d-%m-%Y"):
        """
        Returns the current date.

        Args:
            format_string (str, optional): The desired format for the date.
                                           Defaults to "%d-%m-%Y" (e.g., "27-10-2025").

        Returns:
            str: The current date formatted as a string.
        """
        return datetime.date.today().strftime(format_string)

    def get_current_time(self, format_string="%H:%M:%S"):
        """
        Returns the current time.

        Args:
            format_string (str, optional): The desired format for the time.
                                           Defaults to "%H:%M:%S" (e.g., "14:30:45").

        Returns:
            str: The current time formatted as a string.
        """
        return datetime.datetime.now().strftime(format_string)

    def get_days_between_dates(self, start_date_str, end_date_str, exclude_weekends=False, date_format="%d-%m-%Y"):
        """
        Calculates the number of days between two dates.

        Args:
            start_date_str (str): The start date as a string (e.g., "20-10-2025").
            end_date_str (str): The end date as a string (e.g., "30-10-2025").
            exclude_weekends (bool, optional): If True, weekends (Saturday and Sunday)
                                               are excluded from the count. Defaults to False.
            date_format (str, optional): The format of the input date strings.
                                         Defaults to "%d-%m-%Y".

        Returns:
            int: The number of days between the two dates. Returns -1 if dates are invalid
                 or start date is after end date.
        """
        try:
            start_date = datetime.datetime.strptime(start_date_str, date_format).date()
            end_date = datetime.datetime.strptime(end_date_str, date_format).date()
        except ValueError:
            print(f"Error: Invalid date format. Please use '{date_format}'.")
            return -1

        if start_date > end_date:
            print("Error: Start date cannot be after end date.")
            return -1

        if not exclude_weekends:
            return (end_date - start_date).days
        else:
            delta = end_date - start_date
            business_days = 0
            for i in range(delta.days + 1):  # +1 to include the end date
                current_day = start_date + datetime.timedelta(days=i)
                # weekday() returns 0 for Monday, 6 for Sunday
                if current_day.weekday() < 5:  # Monday to Friday
                    business_days += 1
            return business_days
