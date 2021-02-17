import pandas as pd
import os

cwd = os.getcwd()

users = pd.read_excel('./data/users.xlsx')
users_dir = os.getcwd() + '/data/users'

columns = ['Albümin/Kreatinin (Spot idrar)', 'Demir (Serum/Plazma)',
		   'Ferritin (Serum/Plazma)',
		   'Glike hemoglobin (Hb A1c) (HPLC)', 'HbA1C',
		   'Hemoglobin (HGB) (Hemogram(Tam Kan))',
		   'Kreatinin (Kreatinin (Serum/Plazma))', 'Kreatinin (Spot idrar)',
		   'UIBC']

def clean_invalid_chars(s):
	try:
		s = s.replace(',', '.')
		i = float(s)
		return i
	except ValueError as verr:
		return 0
	except Exception as ex:
		return 0

if __name__ == '__main__':
	files_to_read = []
	data_frames = []
	# will check all files in the dir
	for filename in os.listdir(users_dir):
		if filename.endswith(".xlsx"):
			files_to_read.append(os.path.join(users_dir, filename))
		else:
			continue
	print(files_to_read)
	for d in files_to_read:
		chunks = d.split('/')
		filename = chunks[7]
		user_data = pd.read_excel(d)
		df = user_data[['Test Adı', 'Sonuç']]
		df = df.fillna(0)
		df['Sonuç'] = df['Sonuç'].apply(lambda x: clean(x))
		df['Sonuç'] = df['Sonuç'].astype(float)
		df_mean = df.groupby('Test Adı', as_index=True).mean()
		d = df_mean.T
		user_value = d[d.columns.intersection(columns)]
		f = user_value.rename(columns={'Test Adı': 'Isim Soyisim'}, index={'Sonuç': filename.split('.xlsx')[0]})
		data_frames.append(f)

	main_frame = pd.concat(data_frames)
	print(main_frame)
	main_frame.to_excel("yeni.xlsx")
