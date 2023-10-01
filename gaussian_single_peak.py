#Run this script in a folder and it will convert all MassLynx extracted .txt data files and create an xlsx sheet
#Also plots the mobilligram with Nagy lab styles

#Edited to also fit a Gaussian to the top 50% of the peak and plot it

###COLOR CHOICE
#options are red, blue, green, yellow, orange, purple
from __future__ import print_function
color_choice = 'blue'

#Iterate the .txt -> Excel code through every .txt file in a directory

import os

from scipy.optimize import curve_fit
import matplotlib.pyplot as mpl

import numpy as np
from astropy import modeling

import matplotlib.pyplot as mpl

#Set the path of this script as the directory to look for files in
dir_path = os.path.dirname(os.path.realpath(__file__))


#Finds every .txt file in the directory and adds it to a list
file_list = []

import glob, os
os.chdir(dir_path)
for file in glob.glob("*.txt"):
    file_list.append(file)


for i in file_list:

    import pandas as pd
    import matplotlib.pyplot as plt
    import xlsxwriter
    from pandas import ExcelWriter

    IMS_filename = (str(i))

    pd.options.display.max_rows = 200

    #IMS_filename = "a14_5m.txt"

    data = pd.read_csv(IMS_filename, sep = '\t', header=None, usecols=[0,1])

    data.columns = ['Raw Arrival Time', 'Raw Intensity']

    #Subtract 10 ms form arrival time data and create new column
    data['Arrival Time'] = data['Raw Arrival Time'] - 10

    #Normalize Intensity

    data['Normalized Intensity'] = (data['Raw Intensity'] - data['Raw Intensity'].min())/(data['Raw Intensity'].max()-data['Raw Intensity'].min())

  ################################## GAUSSIAN

    top50_x_precut = data['Arrival Time']
    top50_y_precut = data['Normalized Intensity']

    top50_y = []
    top50_x = []

    for k in range(0, len(top50_x_precut)):
        if top50_y_precut[k] > 0.5:
            top50_x.append(top50_x_precut[k])
            top50_y.append(top50_y_precut[k])
        #if k > 0.5:
           #top50_y.append(k)
        #print(data['Arrival Time'])
    #print(top50_y)

    data_gaus = pd.DataFrame(top50_x, columns = ['Top 50% Arrival Time'])
    data_gaus['Top 50% Intensity'] = top50_y

    #print(data['Normalized Intensity'][10])

  #Define the Gaussian function
    def gauss(x, a, x0, sigma):
        return (a*np.exp(-(x-x0)**2/(2*sigma**2)))


    a = 1
    x0 = sum(top50_x)/len(top50_x)
    sigma = -.6

    guess = np.array([a, x0, sigma])

    #print(guess)

    # Executing curve_fit on noisy data
    popt, pcov = curve_fit(gauss, top50_x, top50_y, guess)



    fig = mpl.figure()
    ax = fig.add_subplot(111)
    ax.plot(top50_x, top50_y, c='k', label='Function')
    ax.scatter(top50_x, top50_y)

    ym = gauss(top50_x, popt[0], popt[1], popt[2])

    data_gaus['Gaussian Fit'] = (popt[0]*np.exp(-(top50_x-popt[1])**2/(2*popt[2]**2)))
    data['Gaussian Fit'] = (popt[0]*np.exp(-(data['Arrival Time']-popt[1])**2/(2*popt[2]**2)))

    data_opt = pd.DataFrame(['a', 'x0', 'sigma'], columns = ['Parameter'])
    data_opt['Value'] = popt






###################################### END of GAUSSIAN

    #Switches .txt for .xlsx (by creating a new variable)
    New_filename = IMS_filename[:-4] + '.xlsx'

    #Sets up the list of colors

    import random

    color_list = ['#FF0000','#009900', '#0000CC', '#9900CC', '#CC9900', '#FF6600' ]
    color_name = ['red', 'green', 'blue', 'purple', 'yellow', 'orange']

    color_dict = dict(zip(color_name, color_list))

    random_color = (color_dict.get(random.choice(color_name)))
    set_color = color_dict.get(color_choice)



    #Writes to an excel sheet
    writer = pd.ExcelWriter(New_filename, engine='xlsxwriter')

    data.to_excel(writer, sheet_name='Sheet1')

    data_gaus.to_excel(writer, sheet_name='Gauss_Fit')

    data_opt.to_excel(writer, sheet_name='Opt')

    worksheet_gaus = writer.sheets['Gauss_Fit']


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

