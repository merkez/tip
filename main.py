import pandas as pd
import os

cwd = os.getcwd()
users = pd.read_excel('./data/users.xlsx')
users_dir = os.getcwd() + '/data/users'
data_file = 'all_data.xlsx'
suffix = '.xlsx'

columns = ['Albümin/Kreatinin (Spot idrar)', 'Demir (Serum/Plazma)',
		   'Ferritin (Serum/Plazma)',
		   'Glike hemoglobin (Hb A1c) (HPLC)', 'HbA1C',
		   'Hemoglobin (HGB) (Hemogram(Tam Kan))',
		   'Kreatinin (Kreatinin (Serum/Plazma))', 'Kreatinin (Spot idrar)',
		   'UIBC']


def clean_invalid_chars(s):
	"""

	:param s:
	char to clean
	:return:
	Returns converted float value if it is possible to convert
	"""
	try:
		s = s.replace(',', '.')
		i = float(s)
		return i
	except ValueError:
		return 0
	except Exception:
		return 0


def get_file_names(directory):
	"""
	:param directory:
	dir to look at for files
	:return:
	Returns list of file names
	"""
	files_to_read = []
	for filename in os.listdir(directory):
		if filename.endswith(suffix):
			files_to_read.append(os.path.join(directory, filename))
		else:
			continue
	return files_to_read


def clean_dataframe(df):
	"""

	:param df:
	Dataframe to clean
	:return:
	Dataframe which is cleaned for the purpose
	"""
	df = df[['Test Adı', 'Sonuç']]
	df = df.fillna(0)
	df['Sonuç'] = df['Sonuç'].apply(lambda x: clean_invalid_chars(x))
	df['Sonuç'] = df['Sonuç'].astype(float)
	return df


def get_average(df):
	"""

	:param df:
	:return:
	Returns mean of given dataframe grouped by column name
	"""
	df_mean = df.groupby('Test Adı', as_index=True).mean()
	d = df_mean.T
	user_value = d[d.columns.intersection(columns)]
	f = user_value.rename(columns={'Test Adı': 'Isim Soyisim'}, index={'Sonuç': filename.split(suffix)[0]})
	return f


if __name__ == '__main__':
	files = get_file_names(users_dir)
	data_frames = []

	for d in files:
		chunks = d.split('/')
		filename = chunks[7]
		user_data = pd.read_excel(d)
		f = clean_dataframe(user_data)
		f = get_average(f)
		data_frames.append(f)

	main_frame = pd.concat(data_frames)
	print(main_frame)
	main_frame.to_excel(data_file)

