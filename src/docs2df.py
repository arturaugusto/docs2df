# coding: utf8
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import pandas
from difflib import SequenceMatcher

class DocxDataFrames(object):
  """docstring for DocxDataFrames"""
  def __init__(self, doc, concat_when_gap_below = None, preprocess_fun = None, col_normalization_mapping = None):
    super(DocxDataFrames, self).__init__()
    self.parent = doc
    self.concat_when_gap_below = concat_when_gap_below
    self.preprocess_fun = preprocess_fun
    self.col_normalization_mapping = col_normalization_mapping

  def iter_block_items(self):
    """
    Yield each paragraph and table child within *parent*, in document order.
    Each returned value is an instance of either Table or Paragraph. *parent*
    would most commonly be a reference to a main Document object, but
    also works for a _Cell object, which itself can contain paragraphs and tables.
    Reference: https://github.com/python-openxml/python-docx/issues/40#issuecomment-90710401
    """
    if isinstance(self.parent, Document):
      parent_elm = self.parent.element.body
    elif isinstance(self.parent, _Cell):
      parent_elm = self.parent._tc
    else:
      raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
      if isinstance(child, CT_P):
        yield Paragraph(child, self.parent)
      elif isinstance(child, CT_Tbl):
        yield Table(child, self.parent)

  def table_with_prev_and_next_block(self):
    data = []
    table_counter = 0
    block_accumulator = []

    for block in self.iter_block_items():

      curr_block_is_table = isinstance(block, Table)
      
      if curr_block_is_table:
        table_counter += 1
      else:
        if len(data) > 0:
          data[table_counter - 1]['next'].append(block)

      if curr_block_is_table:
        data_entry = dict()
        data_entry['table'] = block
        data_entry['prev'] = block_accumulator
        data_entry['next'] = []

        block_accumulator = []

        data.append(data_entry)
      else:
        block_accumulator.append(block)
    return data

  def parse_row(self, row_content, col_tags, secundary_tags):
    values, tags = self.get_values_with_tags_from_row_content(row_content)
    
    if len(secundary_tags) == 0:
      secundary_tags = [''] * len(values)

    if  self.preprocess_fun != None:
      preprocess_data = zip(values, col_tags, secundary_tags)
      values = (map(
        lambda x: self.preprocess_fun(
          **{'value': x[0], 'col_tag': x[1], 'secundary_tag': x[2], 'row_tags': tags}),
          zip(values, col_tags, secundary_tags)
        ))
    return values

  def normalize_cols_text(self, col_tags):
    if self.col_normalization_mapping == None:
      return col_tags

    res = []
    for k in col_tags:
      if k in self.col_normalization_mapping:
        res.append(self.col_normalization_mapping[k])
      else:
        res.append(k)
    return res

  def parse_table(self, table):
    rows = table.rows
    first_row = rows[0]

    col_tags = []
    if not self.is_row_mainly_numeric(first_row):
      col_tags = self.get_row_content(first_row)

    col_tags = self.normalize_cols_text(col_tags)

    content_rows = self.get_content_rows(rows)
    #TODO: optimize this
    tbl_rows_values_with_tags = []
    secundary_tags = []
    for row in content_rows:
      row_content = self.get_row_content(row)
      if not self.is_row_mainly_numeric(row):
        secundary_tags = row_content
      else:
        values = self.parse_row(row_content, col_tags, secundary_tags)
        
        tbl_rows_values_with_tags.append(values)
    
    df = pandas.DataFrame(tbl_rows_values_with_tags, columns = col_tags)
    return df

  def get_values_with_tags_from_row_content(self, row_content):
    row_tags = []
    cell_val_tag_arr = []
    for cell_content in row_content:
      try:
        value = self.txt_to_num(cell_content)
      except Exception:
        row_tags.append(cell_content)
        value = cell_content

      cell_val_tag_arr.append(value)
    return cell_val_tag_arr, row_tags

  def get_content_rows(self, rows):
    return rows[1:]

  def get_row_content(self, row):
    return map(lambda x: x.text, row.cells)

  def txt_to_num(self, txt):
    return float(str(txt).replace(',','.'))

  def is_row_mainly_numeric(self, row, tresh = 0.5):
    cell_n = len(row.cells)
    numeric_cell_counter = 0
    for cell in row.cells:
      try:
        cell_val = self.txt_to_num(cell.text)
        numeric_cell_counter += 1
      except Exception:
        pass
    
    return (numeric_cell_counter / float(cell_n)) > tresh

  def data_blocks_to_readable_arrays(self):
    data = []
    data_blocks = self.table_with_prev_and_next_block()
    for block_grp_dict in data_blocks:
      data_entry = dict()
      data_entry['prev'] = self.join_arrays_to_string(
        map(
          lambda x: x.text, block_grp_dict['prev']
          )
        )
      data_entry['next'] = self.join_arrays_to_string(
        map(
          lambda x: x.text, block_grp_dict['next']
          )
        )
      data_entry['table'] = self.parse_table(block_grp_dict['table'])
      data.append(data_entry)
    return data

  def join_arrays_to_string(self, blocks):
    return ''.join(map(lambda x: x.strip()+'\n', blocks))

  def concat_table_data_with_small_gap(self, df, concat_when_gap_below = 50):
    res = []
    for i, row in enumerate(df):
      if (len(row['prev']) < concat_when_gap_below) and i > 0:
        res[-1]['table'] = pandas.concat([res[-1]['table'], row['table']], ignore_index=True)
      else:
        res.append(row)
    return res

  def get_dataframes(self):
    self.table_with_prev_and_next_block()
    df = self.data_blocks_to_readable_arrays()
    if self.concat_when_gap_below != None:
      df = self.concat_table_data_with_small_gap(df, self.concat_when_gap_below)
    return df

class AggregatedDocxDataFrame(object):
  """docstring for AggregatedDocxDataFrame"""
  def __init__(self, docxDataFrames):
    super(AggregatedDocxDataFrame, self).__init__()
    self.docxdf_list = list()
    for docxdf in docxDataFrames:
      assert isinstance(docxdf, DocxDataFrames)
      self.docxdf_list.append(docxdf)

  def similar(self, a, b):
    return SequenceMatcher(None, a, b).ratio()

  def default_roi_fun(self, prv, nxt, df):
    return prv

  def get_similar_tables(self, query_text, tresh = 0.7, roi_fun = None):
    res = []
    for n, docxdf in enumerate(self.docxdf_list):
      best_match_table = None
      best_match_similarity = 0
      for t in docxdf.get_dataframes():
        if roi_fun != None:
          text_to_compare = roi_fun(t['prev'], t['next'], t['table'])
        else:
          text_to_compare = self.default_roi_fun(t['prev'], t['next'], t['table'])
        similarity = self.similar(text_to_compare, query_text)
        if ( (similarity > tresh) and (similarity > best_match_similarity) ):
          best_match_table = t['table']
          best_match_similarity = similarity
      res.append(dict({'similarity': best_match_similarity, 'table': best_match_table}))
    return res