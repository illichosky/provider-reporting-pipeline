import requests
import time
from settings import *

def poll_job(s, redash_url, job):
    # TODO: add timeout
    while job['status'] not in (3, 4):
        response = s.get('{}/api/jobs/{}'.format(redash_url, job['id']))
        job = response.json()['job']
        time.sleep(1)

    if job['status'] == 3:
        return job['query_result_id']

    return None

def get_fresh_query_result(redash_url, query_id, params):
    s = requests.Session()
    s.headers.update({'Authorization': 'Key {}'.format(REDASH_USER_API_KEY)})

    response = s.post('{}/api/queries/{}/refresh'.format(redash_url, query_id), params=params)

    if response.status_code != 200:
        raise Exception('Refresh failed. '+str(response.status_code))

    result_id = poll_job(s, redash_url, response.json()['job'])

    if result_id:
        response = s.get('{}/api/queries/{}/results/{}.json'.format(redash_url, query_id, result_id))
        if response.status_code != 200:
            raise Exception('Failed getting results.')
    else:
        raise Exception('Query execution failed.')

    return response.json()['query_result']['data']['rows']
