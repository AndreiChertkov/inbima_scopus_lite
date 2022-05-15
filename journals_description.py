def _color_q(v):
    if v == 'Q1': return '00FF6600'
    if v == 'Q2': return '00FFCC00'
    if v == 'Q3': return '00808000'
    if v == 'Q4': return '00666699'


description = {
    'journals': {
        '#opts': {
            'author': 'Andrei Chertkov',
            'sheet_password': None,
            'freeze': 'B2',
            'height': {
                'head': 40,
                'data': 80,
            },
            'font': {
                'size': {
                    'head': 16,
                    'data': 12,
                },
                'bold': {
                    'head': True,
                    'data': False,
                },
                'color': {
                    'head': '00000000',
                    'data': '00000000',
                    'error': '00800000',
                },
            },
            'align': {
                'hor': {
                    'head': 'center',
                    'data': 'left',
                },
                'ver': {
                    'head': 'center',
                    'data': 'center',
                },
            },
            'wrap': {
                'head': False,
                'data': True,
            },
            'fill': {
                'head': '00C0C0C0',
                'data': None,
                'empty': '00808080',
                'tab': '00008080',
            },
            'border': {
                'left': {
                    'head': False,
                    'data': False,
                },
                'top': {
                    'head': False,
                    'data': False,
                },
                'right': {
                    'head': False,
                    'data': False,
                },
                'bottom': {
                    'head': True,
                    'data': True,
                },
            },
            'note': {
                'width': 500,
                'height': 300,
            }
        },
        'title': {
            'name': 'Title',
            'size': 30,
        },
        'issn': {
            'name': 'ISSN',
            'size': 12,
        },
        'eissn': {
            'name': 'EISSN',
            'size': 12,
            'separator': True,
        },
        'score': {
            'name': 'Score',
            'note': 'Scopus Cite Score index (the number of citations of papers published in the journal in four years is divided by the number of peer-reviewed papers indexed in Scopus and published in the same four years).',
            'size': 12,
        },
        'cited': {
            'name': 'Cited',
            'size': 12,
            'separator': True,
        },
        'quartile': {
            'name': 'Quartile',
            'size': 17,
            'note': 'Best quartile of the journal according to Scopus.',
            'align': 'center',
            'color_func': _color_q,
        },
        'q1': {
            'name': 'Q1',
            'note': 'Quartile Q1 fields for the journal according to Scopus.',
            'size': 30,
            'join': '; '
        },
        'q2': {
            'name': 'Q2',
            'note': 'Quartile Q2 fields for the journal according to Scopus.',
            'size': 30,
            'join': '; '
        },
        'q3': {
            'name': 'Q3',
            'note': 'Quartile Q3 fields for the journal according to Scopus.',
            'size': 30,
            'join': '; '
        },
        'q4': {
            'name': 'Q4',
            'note': 'Quartile Q4 fields for the journal according to Scopus.',
            'size': 30,
            'join': '; '
        },
        'q0': {
            'name': 'Q None',
            'note': 'Quartile fields without quartile value for the journal.',
            'size': 30,
            'join': '; ',
            'separator': True,
        },
        'snip': {
            'name': 'SNIP',
            'note': 'Source Normalized Impact per Paper. It measures contextual citation impact by weighting citations based on the total number of citations in a subject field, using Scopus data.',
            'size': 12,
        },
        'sjr': {
            'name': 'SJR',
            'note': 'SCImago Journal Rank.',
            'size': 12,
        },
        'top_10': {
            'name': 'In top 10',
            'note': 'Is the journal in Scopus top 10?',
            'size': 12,
        },
        'id': {
            'name': 'Id',
            'note': 'Id in the Scopus database.',
            'size': 10,
        },
        'url': {
            'name': 'Url',
            'note': 'Journal page in the Scopus database.',
            'size': 12,
            'separator': True,
        },
        'open': {
            'name': 'Open',
            'note': 'Is it open access journal?',
            'size': 12,
        },
        'publisher': {
            'name': 'Publisher',
            'size': 20,
        },
        'type': {
            'name': 'Type',
            'note': 'Type of the journal according to Scopus.',
            'size': 10,
            'separator': True,
        },
        'error': {
            'name': 'Error',
            'note': 'Description of errors that occurred during automatic parsing of the journal.',
            'size': 25,
            'separator': True,
        },
    }
}
