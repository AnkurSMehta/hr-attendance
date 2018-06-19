from datetime import date, timedelta

def gen_all_sundays(year):
    results = []

    first_day_of_year = date(year, 1, 1)
    first_day_of_year += timedelta(days = 6 - first_day_of_year.weekday())  # First Sunday

    while first_day_of_year.year == year:
      results.append(str(first_day_of_year))
      first_day_of_year += timedelta(days = 7)
    
    return results

if __name__ == "__main__":
    print(gen_all_sundays(2018))