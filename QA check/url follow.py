import requests

def init():
    console = input("Type the URL: ")
    get_status_code_from_request_url(console)


def get_status_code_from_request_url(url, do_restart=True):
    try:
        r = requests.get(url)
        if len(r.history) < 1:
            print("Status Code: " + str(r.status_code))
        else:
            print("Status Code: 301. Below are the redirects")
            h = r.history
            i = 0
            for resp in h:
                print("  " + str(i) + " - URL " + resp.url + " \n")
                i += 1
        if do_restart:
            init()
    except requests.exceptions.MissingSchema:
        print("You forgot the protocol. http://, https://, ftp://")
    except requests.exceptions.ConnectionError:
        print("Sorry, but I couldn't connect. There was a connection problem.")
    except requests.exceptions.Timeout:
        print("Sorry, but I couldn't connect. I timed out.")
    except requests.exceptions.TooManyRedirects:
        print("There were too many redirects.  I can't count that high.")


init()