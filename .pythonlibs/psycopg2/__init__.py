"""Pure-Python psycopg2 compatibility layer backed by pg8000."""

from pg8000 import dbapi


class _Cursor:
    def __init__(self, cursor):
        self._cursor = cursor

    def __getattr__(self, name):
        return getattr(self._cursor, name)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self._cursor.close()
        return False


class _Connection:
    def __init__(self, connection):
        self._connection = connection

    def __getattr__(self, name):
        return getattr(self._connection, name)

    def cursor(self, *args, **kwargs):
        return _Cursor(self._connection.cursor(*args, **kwargs))


def connect(*args, **kwargs):
    return _Connection(dbapi.connect(*args, **kwargs))


Error = dbapi.Error
DatabaseError = dbapi.DatabaseError
OperationalError = dbapi.OperationalError
