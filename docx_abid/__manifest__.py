# -*- coding: utf-8 -*-

{
    'name': 'DOCX Report',
    'description': 'Design reports using Microsoft Word. Export Odoo data to DOCX and PDF.',
    'summary': 'Export Odoo data to Microsoft Office — DOCX & PDF reports',
    'category': 'All',
    'version': '19.0.1.0.0',
    "license": "OPL-1",
    'author': 'Abdulkaraim Osman',
    'depends': [
        'base', 'web'
    ],
    'installable': True,
    'application': True,
    "external_dependencies": {
        "python": ["pybase64"],
        "bin": ["unoconv"],
    },
    'data': [
        'data/templates.xml',
        'report.xml',
        'views/report_view.xml'
    ],

    'images': ['static/description/icon.png', 'static/description/banner.png', 'static/description/logo.png'],
    'auto_install': False,
    'assets': {
        'web.assets_backend': [
            'docx_abid/static/src/scss/theme_screenshot.scss',
        ]
    }
}
