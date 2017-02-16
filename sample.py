import docx
from src.docs2df import *
import numpy as np
import matplotlib.pyplot as plt

docs_names = ['test.docx','test_2.docx','test_3.docx']

col_normalization_mapping = {
  'A.A.': 'AA',
  'B.B.': 'BB',
  'C.C.': 'CC',
  'D.D.': 'DD',
  'E.E.': 'EE',
  'F.F.': 'FF',
  'G.G.': 'GG',
}

def preprocess_fun(**args):
  if 'mV' in args['secundary_tag']:
    return args['value'] / 1000.0
  return args['value']

docxdf_list = []
for doc_name in docs_names:
  doc = docx.Document('tests/{}'.format(doc_name))
  docxdf = DocxDataFrames(doc)
  docxdf.concat_when_gap_below = 50
  docxdf.col_normalization_mapping = col_normalization_mapping
  docxdf.preprocess_fun = preprocess_fun
  docxdf_list.append(docxdf)

aggr = AggregatedDocxDataFrame(docxdf_list)

query_text = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent condimentum nisl nibh. Nulla non turpis odio. Integer et mollis eros, vitae varius enim. Sed eget ex pharetra, sodales lorem nec, iaculis urna. Nunc scelerisque sollicitudin consequat. Sed ut nunc a odio malesuada convallis. Sed facilisis suscipit tristique. Etiam nec tellus eget mi blandit vehicula. Donec finibus magna sit amet sodales posuere. Integer ante dui, pellentesque eu arcu vitae, dignissim tristique justo. Nullam mollis posuere dictum. Pellentesque id mi at purus ullamcorper volutpat. Proin euismod nisl odio, in lobortis nulla ullamcorper ut. Nam ac purus quis ligula pharetra imperdiet.'
    
res = aggr.get_similar_tables(query_text)

data = map(lambda x: x['table'].loc[lambda df: (df.AA < 0.051) & (df.AA > 0.049) & (df.XX == '60 MHz'), :], res)

data = list(filter(lambda x: not x.empty, data))

def get_data_list(key):
  return map(lambda x:float(x[key]), data)

y = get_data_list('CC')
x = np.arange(len(y))
yerr = get_data_list('EE')


# First illustrate basic pyplot interface, using defaults where possible.
plt.figure()
plt.errorbar(x, y, yerr=[map(lambda d: d[0], zip(yerr, y)), map(lambda d: d[0], zip(yerr, y))], fmt='--o')
plt.title("Errorbars")
plt.margins(0.05)

plt.show()



