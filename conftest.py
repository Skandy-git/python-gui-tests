import pytest
from fixture.application import Application


@pytest.fixture(scope="session")
def app(request):
    fixture = Application("C:\\Users\\User\\Desktop\\NIKITA\\addressbook\\AddressBook.exe")
    request.addfinalizer(fixture.destroy)
    return fixture