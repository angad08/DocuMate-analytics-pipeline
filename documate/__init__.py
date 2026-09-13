"""
DocuMate.

Every version (v3, X, Y, Z, O) does the same job. They only differ in
three places:

    where records come from   ->  Excel, or the database      (sources/)
    how the Word file is made ->  docxtpl, or Word Mail Merge (engines/)
    where PRINTED is written  ->  back to Excel, or the database

So there is one copy of each of those, and the version just says which
combination to use. The list of versions is in versions/registry.py.
"""

__version__ = "4.0.0"
