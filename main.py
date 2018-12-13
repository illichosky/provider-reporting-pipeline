import pandas as pd
import json
from redash import get_fresh_query_result


# TODO Create unit tests

def get_reports(report_list, params):
    """Get reports from Redash

    This function iterates through the report_list, creates an object containing each report and its title, and
    append it to a list, that will be returned in the end of the execution.

    :param report_list: List of objects contaning the desired title of the query and the query_id to be queried.
        from redash. Example:
        [
            {'title': 'Individual ratings', 'query_id': 140},
            {'title': 'NPUs per provider', 'query_id': 145}
        ]

    :param params: It's a dict contaning the query parameters.
        Example: {'p_reservation_type': 'talixo', 'p_start_date': '2018-11-01 00:00', 'p_end_date': '2018-11-30 23:59'}

    :return:
        Returns a list of objects containing the json report and its title.

        Example:
        {
            [
                {report: {json}, title: report_title}
            ]
        }
    """

    output = []
    for report in report_list:
        result = json.dumps(get_fresh_query_result('https://data.mozio.com', report['query_id'], params))
        output.append({'report': result, 'title': report['title']})
    return output


def generate_excel(reports):
    """Generates the Excel file of the reports

    It gets a list of reports and transforms into a single Excel file, creating a sheet for each report in the list.

    :param reports: Uses the object returned from get_reports() as parameter for this function.
        Example:
        {
            [
                {report: {json}, title: report_title}
            ]
        }

    :return: Outputs an .xlsx file in the root directory
    """

    writer = pd.ExcelWriter('output.xlsx')
    for sheet in reports:
        pd.read_json(sheet['report']).to_excel(writer, sheet_name=sheet['title'], index=False)
    writer.save()


if __name__ == '__main__':
    report_list =  [
            {'title': 'Individual ratings', 'query_id': 140},
            {'title': 'NPUs per provider', 'query_id': 145},
            {'title': 'Ratings per provider', 'query_id': 144},
            {'title': 'Individual issues', 'query_id':139},
            {'title': 'Ratings per airport', 'query_id': 142},
            {'title': 'NPUs per airport', 'query_id': 138}
                 ]

    start_date = '2018-11-01 00:00'
    end_date = '2018-11-30 23:59'
    reservation_type = 'talixo'
    params = {'p_reservation_type': reservation_type, 'p_start_date': start_date, 'p_end_date': end_date}

    # Should I use keyword arguments to pass each parameter and populate the params object inside get_reports()?
    # It worth noting the params change for other reports - I'm not sure what's best way to do it...
    generate_excel(get_reports(report_list, params))