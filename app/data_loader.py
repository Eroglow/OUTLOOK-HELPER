from typing import List, Dict, Tuple
import pandas as pd


def read_tabular_file(file_path: str) -> Tuple[pd.DataFrame, List[str]]:
	"""
	Read .xlsx or .csv into a DataFrame with best-effort encoding sniffing for CSV.
	Returns (df, columns)
	"""
	if file_path.lower().endswith(".xlsx") or file_path.lower().endswith(".xlsm"):
		df = pd.read_excel(file_path)
	else:
		# CSV: try default, fallback encodings
		try:
			df = pd.read_csv(file_path)
		except UnicodeDecodeError:
			for enc in ["utf-8-sig", "utf-16", "cp950", "big5", "latin1"]:
				try:
					df = pd.read_csv(file_path, encoding=enc)
					break
				except UnicodeDecodeError:
					continue
			else:
				raise
	columns = [str(c) for c in df.columns]
	return df, columns


def read_from_clipboard() -> Tuple[pd.DataFrame, List[str]]:
	"""Read a table from the clipboard (e.g., copied from Excel)."""
	df = pd.read_clipboard(sep="\t")
	columns = [str(c) for c in df.columns]
	return df, columns


def dataframe_to_records(df: pd.DataFrame) -> List[Dict[str, object]]:
	"""Convert DataFrame rows to a list of dict records with string keys."""
	records: List[Dict[str, object]] = []
	for _, row in df.iterrows():
		record: Dict[str, object] = {}
		for col in df.columns:
			record[str(col)] = row[col]
		records.append(record)
	return records
