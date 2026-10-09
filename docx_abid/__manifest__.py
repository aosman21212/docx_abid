# -*- coding: utf-8 -*-
# =============================================================================
#  DOCX Report
# -----------------------------------------------------------------------------
#  Location  : King Abdulaziz Branch Road, Riyadh, Saudi Arabia
#  Email     : sales@leapai.ai
#  Phone     : +966 53 553 3627
#  Website   : https://leapai.ai
#  Developer : Abdulkaraim Osman — Tech Manager | Backend Engineer | DevOps Engineer
#              at Bab International Corp For Specialized Services
#  LinkedIn  : https://www.linkedin.com/in/abdulkaraim-o-385b7a110/
# =============================================================================

{
    'name': 'DOCX Report',
    'description': 'Design reports using Microsoft Word. Export Odoo data to DOCX and PDF.',
    'summary': 'Export Odoo data to Microsoft Office — DOCX & PDF reports',
    'category': 'All',
    'version': '19.0.1.0.0',
    "license": "OPL-1",
    'author': 'LeapAI',
    'maintainer': 'Abdulkaraim Osman',
    'support': 'sales@leapai.ai',
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
