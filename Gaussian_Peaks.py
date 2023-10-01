#Run this script in a folder and it will convert all MassLynx extracted .txt data files and create an xlsx sheet
#Also plots the mobilligram with Nagy lab styles

#Edited to also fit a Gaussian to the top 50% of the peak and plot it

###COLOR CHOICE
#options are red, blue, green, yellow, orange, purple

from __future__ import print_function
import os
from scipy.signal import find_peaks
from scipy.optimize import curve_fit
import numpy as np
import matplotlib.pyplot as mpl
import glob, os
import random
import pandas as pd
import matplotlib.pyplot as plt



def main():

    #Iterate the .txt -> Excel code through every .txt file in a directory

    #Set the path of this script as the directory to look for files in
    dir_path = os.path.dirname(os.path.realpath(__file__))

    #Finds every .txt file in the directory and adds it to a list

    file_list = txt_finder(dir_path)

    opt_params = []

    # Converts txt file to pandas df

    for i in range(len(file_list)):

        data = txt_to_df(file_list[i])


        ################################## GAUSSIAN
    
        data_gauss = alldata(data)

        opt_found, opt_fit = fit_gauss(data_gauss)
        
        opt_params.append(opt_found)
        

        opt_df = pd.DataFrame(opt_found).T


###################################### END of GAUSSIAN


        # Write to Excel sheet
        data['Gaussian Fit'] = opt_fit

        xls_file_out(file_list[i], data, opt_df)

    # CSV store fitted parameters
    cwd = os.getcwd()
    file = cwd + "/opt_params.csv"
    
    params_df = pd.DataFrame(opt_params)
    params_df.to_csv(file)


# Finds every .txt file in the directory and adds it to a list
def txt_finder(dir_path):
    file_list = []


    os.chdir(dir_path)
    for file in glob.glob("*.txt"):
        file_list.append(file)

    return file_list

# Converts IMS txt file to pandas df
def txt_to_df(path_to_file):
        
    IMS_filename = str(path_to_file)

    pd.options.display.max_rows = 200


    data = pd.read_csv(IMS_filename, sep = '\t', header=None, usecols=[0,1])

    data.columns = ['Raw Arrival Time', 'Raw Intensity']

    #Subtract 10 ms form arrival time data and create new column
    data['Arrival Time'] = data['Raw Arrival Time'] - 10
    data['Normalized Intensity'] = (data['Raw Intensity'] - data['Raw Intensity'].min())/(data['Raw Intensity'].max()-data['Raw Intensity'].min())

    return data

# Use all data

def alldata(data):
    x = data['Arrival Time']
    y = data['Normalized Intensity']

    data_gaus = pd.DataFrame(x, columns = ['Arrival Time'])
    data_gaus['Normalized Intensity'] = y

    return data_gaus

# Use only data over 50% intensity (optional)

def top50(data):

    top50_x_precut = data['Arrival Time']
    top50_y_precut = data['Normalized Intensity']

    top50_y = []
    top50_x = []

    for i in range(0, len(top50_x_precut)):
        if top50_y_precut[k] > 0.5:
            top50_x.append(top50_x_precut[i])
            top50_y.append(top50_y_precut[i])


    data_gaus = pd.DataFrame(top50_x, columns = ['Top 50% Arrival Time'])
    data_gaus['Top 50% Intensity'] = top50_y

    return data_gaus


# Function to find the number and position of each peak. Prominence value toggles sensitivity
def peak_list(data):
    if type(data) == str and os.path.isfile(data):

        data = pd.read_csv(data, sep='\t').to_numpy()

    x = data[:, 0]

    y = data[:, 1]
    norm_y = (y - np.min(y))/ np.max(y)-np.min(y)

    peaks, _ = find_peaks(norm_y, prominence=.01)
    spacing = round(x[1]- x[0], 6)
    rescale_x = ((peaks, norm_y[peaks])[0] * spacing) + np.min(x)
    rescale_peaks = pd.DataFrame((rescale_x, (peaks, norm_y[peaks])[1]))

    return(rescale_peaks)

# Function to fit multiple gaussians to data
def fit_gauss(data_gauss):
    
    numpy = data_gauss.to_numpy()
    peaks = peak_list(numpy)

    x = (numpy[:, 0])
    y = (numpy[:, 1])


    # Use found peaks as x0 guesses
    x0_guess = peaks.loc[0,0]
    # Base width guess on arrival time
    width_guess = np.average(x)/200
    guess = [x0_guess, 2, width_guess]
    for i in range(len(peaks.loc[0, :])-1):
        guess += [peaks.loc[0,1+i], 2, width_guess]

    popt, pcov = curve_fit(func, x, y, p0=guess)
    print(popt)
    fit = func(x, *popt)

    plt.plot(x, y)
    plt.plot(x, fit , 'r-')
    plt.show()


    #    data_opt = pd.DataFrame(['a', 'x0', 'sigma'], columns = ['Parameter'])
    #    data_opt['Value'] = popt
        
    return popt, fit

# Define the Gaussian function
def func(x, *params):
    y = np.zeros_like(x)
    for i in range(0, len(params), 3):
        ctr = params[i]
        amp = params[i+1]
        wid = params[i+2]
        y = y + amp * np.exp( -((x - ctr)/wid)**2)
    return y


def xls_file_out(file_path, data, opt_df):
    
    color_choice = 'blue'
    
    #Switches .txt for .xlsx (by creating a new variable)
    New_filename = (os.path.basename(file_path)).replace('.txt','.xlsx')

    #Sets up the list of colors

    color_list = ['#FF0000','#009900', '#0000CC', '#9900CC', '#CC9900', '#FF6600' ]
    color_name = ['red', 'green', 'blue', 'purple', 'yellow', 'orange']

    color_dict = dict(zip(color_name, color_list))

    random_color = (color_dict.get(random.choice(color_name)))
    set_color = color_dict.get(color_choice)



    #Writes to an excel sheet
    writer = pd.ExcelWriter(New_filename, engine='xlsxwriter')

    data.to_excel(writer, sheet_name='Sheet1')

    # Names columns
    opt_df.columns = [
    'x0 ' + str((i // 3) + 1) if (i + 3) % 3 == 0 else 
    'A ' + str((i // 3) + 1) if (i + 2) % 3 == 0 else 
    'w ' + str((i // 3) + 1) if (i + 1) % 3 == 0 else 
    col for i, col in enumerate(opt_df.columns)
    ]

    opt_df.to_excel(writer, sheet_name='Opt')


    workbook = writer.book
    worksheet = writer.sheets['Sheet1']


    chart = workbook.add_chart({
        'type': 'scatter',
        'subtype': 'smooth',
    })

    # Configure the series of the chart from the dataframe data.

    # Configure the first series.
    chart.add_series({
        'name':       None,
        'categories': '=Sheet1!$D$2:$D$201',
        'values':     '=Sheet1!$F$2:$F$201',
        'border': {'color': random_color}
    })

    # Configure the gaussian series.
    chart.add_series({
        'name':       None,
        'categories': '=Sheet1!$D$2:$D$201',
        'values':     '=Sheet1!$E$2:$E$201',
        'border': {'color': '#D3D3D3'},
        'width': 1,
    })


    # Configure the chart axes.
    chart.set_x_axis({
        'major_gridlines': {'visible': False},
        'minor_gridlines': {'visible': False},
        'major_tick_mark': 'outside',
        'minor_tick_mark': 'outside',
        'minor_unit': 1,
        'name': 'Arrival Time (ms)',
        'min' : int(data['Arrival Time'].min() + 2),
        'max' : int(data['Arrival Time'].max() - 1),
        'name_font' : {'name': 'Arial', 'size': 14, 'bold' : True},
        'num_font' : {'name': 'Arial', 'size': 10, 'bold' : False}
    })
    chart.set_y_axis({'name': 'Relative Intensity',
        'major_gridlines': {'visible': False},
        'minor_gridlines': {'visible': False},
        'major_tick_mark': 'outside',
        'minor_tick_mark': 'outside',
        'minor_unit': 1,
        'min' : 0,
        'max' : 1.05,
        'num_font' : {'name': 'Arial', 'size': 10, 'bold' : False},
        'name_font' : {'name': 'Arial', 'size': 14, 'bold' : True}
    })

    chart.set_plotarea({
        'layout': {
            'x':      0.2,
            'y':      0.0,
            'width':  0.9,
            'height': 0.8,
        }
    })

    #Remove Legend
    chart.set_legend({'none': True})

    #Remove chart border
    chart.set_chartarea({'border': {'none': True}})


    # Insert the chart into the worksheet.
    worksheet.insert_chart('G2', chart, {'x_scale': 0.732, 'y_scale': 1})



    #Saves the excel sheet
    writer.save()

if __name__ == "__main__":
    main()
