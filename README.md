# adotypes

This is a package aiming to comply with [Python DBAPI2.0](https://peps.python.org/pep-0249/), implemented with ADO COM objects as the backend using comtypes.

This package is still a prototype. Only the essential methods such as [`connect`](https://peps.python.org/pep-0249/#connect), [`.close`](https://peps.python.org/pep-0249/#Connection.close), [`commit`](https://peps.python.org/pep-0249/#commit), [`rollback`](https://peps.python.org/pep-0249/#rollback), [`cursor`](https://peps.python.org/pep-0249/#cursor), [`execute`](https://peps.python.org/pep-0249/#id20) and [`fetch`](https://peps.python.org/pep-0249/#fetchone)es are currently implemented. 

Calling methods that have not yet been implemented in [`Connection`](https://peps.python.org/pep-0249/#connection-objects) or [`Cursor`](https://peps.python.org/pep-0249/#cursor-objects) will raise a `NotImplementedError`.

[The error types](https://peps.python.org/pep-0249/#exceptions) may not be appropriate, and [the type objects](https://peps.python.org/pep-0249/#type-objects-and-constructors) are insufficient.

If there are contributors who can help resolve the mentioned issues, we gladly welcome them.
