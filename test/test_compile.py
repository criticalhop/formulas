# from hyperc import solve
# import hyperc.poc_symex
import os.path
import schedula
from formulas.excel import ExcelModel, BOOK
from formulas.excel.xlreader import load_workbook


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