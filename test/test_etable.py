# from hyperc import solve
# import hyperc.poc_symex
from collections import defaultdict
import os.path
import schedula
from formulas.excel import ExcelModel, BOOK, ERR_CIRCULAR
from formulas.excel.xlreader import load_workbook
from formulas.functions import is_number
import formulas
import unidecode
import string

BOOK = schedula.Token('Book')
def _book2dict(book):
    res = {}
    for ws in book.worksheets:
        s = res[ws.title.upper()] = {}
        for k, cell in ws._cells.items():
            value = getattr(cell, 'value', None)
            if value is not None:
                s[cell.coordinate] = value
    return res


def _res2books(res):
    return {k.upper(): _book2dict(v[BOOK]) for k, v in res.items()}


def _file2books(*fpaths):
    d = os.path.dirname(fpaths[0])
    return {os.path.relpath(fp, d).upper(): _book2dict(
        load_workbook(fp, data_only=True)
    ) for fp in fpaths}

def str_to_py(sheet_name: str):
    trans_name = unidecode.unidecode(sheet_name).replace(' ', '_').upper()
    trans_name = list(trans_name)
    for i, character in enumerate(trans_name):
        if character not in (string.ascii_uppercase + string.digits + "_"):
            trans_name[i] = "_"
    return f"{''.join(trans_name)}"

class ETable:
    def __init__(self, filename) -> None:
        self.filename = filename
        
    def calculate(self):
        xl_mdl = ExcelModel()
        xl_mdl.loads(self.filename)
        # for book in xl_mdl.books.values():
        #     for coord in list(book.values())[0].active._cells:
        #         # book.Book.active._cells[coord] = 99
        #         list(book.values())[0].active._cells[coord].value = 99

        # for coord in xl_mdl.cells:
        #     if xl_mdl.cells[coord].value is not schedula.EMPTY:
        #     #    xl_mdl.cells[coord].value = 44
        #     #    xl_mdl.dsp.default_values[coord]['value'] = 66
        #        xl_mdl.dsp.default_values[coord]['value'] = hyperc.poc_symex.HCProxy(
        #            wrapped=xl_mdl.cells[coord].value, name=coord, parent=None, place_id="__STATIC")
        xl_mdl.calculate()
        code = defaultdict(list)

        for node_key, node_val in xl_mdl.dsp.dmap.nodes.items():
            if ('inputs' in node_val) and ('outputs' in node_val):
                out_py = str_to_py(node_val['outputs'][0])
                code[out_py].append(f'def {out_py}():')
                code[out_py].append(f'    #{node_key}')
                for input in node_val.inputs:
                    cell = formulas.Parser().ast("="+list(formulas.Parser().ast(f'={input}')[1].compile().dsp.nodes.keys())[0].replace(" = -","=-"))[0][0].attr
                    letter = cell.c1
                    number = cell.r1
                    sheet_name = str_to_py(cell.name)
                    var_name = f'var_tbl_{sheet_name}__hct_direct_ref__{number}_{letter}'
                    code[out_py].append(f'    {var_name} = HCT_STATIC_OBJECT.{sheet_name}_{number}.{letter}')


 
        # xl_mdl.calculate()
        # # xl_mdl.add_book(self.link_filename)
        # xl_mdl.write(dirpath=os.path.dirname(__file__))
        # # xl_mdl.finish()
        # # xl_mdl.calculate()
        # # xl_mdl.dsp.dispatch()
        # print('Finished excel-model')


        # xl_mdl.calculate({"'[EXTRA.XLSX]EXTRA'!A1:B1": [[1, 1]]})

        # books = _res2books(xl_mdl.write(xl_mdl.books))

        # msg = '%sCompared overwritten results in %.2fs.\n' \
        #         '%sComparing fresh written results.'

        # res_book = _res2books(xl_mdl.write())


def test_etable():
    mydir = os.path.dirname(__file__)
    # file='trucks.xlsx'
    file='summm.xlsx'
    # file = 'plus.xlsx'
    # file = 'plus_selectif.xlsx'

    et = ETable(os.path.join(mydir, file))
    et.calculate()

def test_excel_model_compile():
    mydir = os.path.join(os.path.dirname(__file__), 'test_files')
    _filename = 'test.xlsx'
    _filename_compile = 'excel.xlsx'
    _link_filename = 'test_link.xlsx'
    _filename_circular = 'circular.xlsx'
    filename = os.path.join(mydir, _filename)
    link_filename = os.path.join(mydir, _link_filename)
    filename_compile = os.path.join(mydir, _filename_compile)
    filename_circular = os.path.join(mydir, _filename_circular)

    results = _file2books(filename, link_filename)
    schedula.get_nested_dicts(results, 'EXTRA.XLSX', 'EXTRA').update({
        'A1': 1, 'B1': 1
    })
    results_compile = _book2dict(
        load_workbook(filename_compile, data_only=True)
    )['DATA']
    results_circular = _file2books(filename_circular)
    maxDiff = None

    xl_model = ExcelModel().loads(filename_compile).finish()
    inputs = ["A%d" % i for i in range(2, 5)]
    outputs = ["C%d" % i for i in range(2, 5)]
    func = xl_model.compile(
        ["'[excel.xlsx]DATA'!%s" % i for i in inputs],
        ["'[excel.xlsx]DATA'!%s" % i for i in outputs]
    )
    i = schedula.selector(inputs, results_compile, output_type='list')
    res = schedula.selector(outputs, results_compile, output_type='list')
    # assertEqual([x.value[0, 0] for x in func(*i)], res)
    # assertIsNot(xl_model, copy.deepcopy(xl_model))
    # assertIsNot(func, copy.deepcopy(func))
    xl_model = ExcelModel().loads(filename_circular).finish(circular=1)
    func = xl_model.compile(
        ["'[circular.xlsx]DATA'!A10"],
        ["'[circular.xlsx]DATA'!E10"]
    )
    # assertEqual(func(False).value[0, 0], 2.0)
    # assertIs(func(True).value[0, 0], ERR_CIRCULAR)
    # assertIsNot(xl_model, copy.deepcopy(xl_model))
    # assertIsNot(func, copy.deepcopy(func))