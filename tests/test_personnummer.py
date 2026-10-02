from app.personnummer import find_candidates,valid_date,valid_luhn

def test_candidate(): assert find_candidates('Personnummer 811218-9876')[0][1]=='811218-9876'
def test_date(): assert valid_date('811218-9876')
def test_luhn(): assert valid_luhn('811218-9876')
