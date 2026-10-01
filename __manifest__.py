{
    'name': 'Universal Bank Statement Import',
    'version': '17.0.1.0.0',
    'category': 'Accounting',
    'summary': 'Import any format of CSV/XLSX Bank Statements robustly',
    'depends': ['account'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/import_bank_statement_view.xml',
        'views/account_journal_views.xml',
    ],
    'installable': True,
    'application': True,
    'license': 'LGPL-3',
}
