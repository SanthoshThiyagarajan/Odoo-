{
    'name': 'Sales Return',
    'version': '1.0',
    'category': 'Sales',
    'depends': ['sale_management', 'sale'],
    'data': [
        'security/ir.model.access.csv',
        'views/views.xml',
        'views/views_menu.xml',
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
}
