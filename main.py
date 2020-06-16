# -*- coding: utf-8 -*-
"""
Created on Mon Jun 15 13:41:02 2020

@author: arseniumn
"""
import pandas as pd
import matplotlib.pyplot as plt
from pandas import ExcelWriter
from pandas import ExcelFile

def main():
    # Add the (.xlsx) file in the path e.g. the name of example is simulations.xlsx
    excel_file = pd.ExcelFile('excel-files/example.xlsx')
    
    # Get the number of worksheets that the excel file contains
    number_of_sheets = len(excel_file.sheet_names)
    
    # For each sheet, read the data
    for i in range (number_of_sheets):
        ax = plt.gca()
        
        # Set the X axis Label
        ax.set_xlabel("Iteration")
        
        # Set the Y axis Label
        ax.set_ylabel("Decrease Rate (%)")
        
        # Store the excel sheet in data frame 
        df = pd.read_excel('excel-files/example.xlsx', 'Sheet'+str(i+1))
        
        # Declare what kind of plot we want
        # Types that you can insert instead of 'line'
        # line, bar, barh, hist, box, kde, density, area, pie, scatter, hexbin
        # X and Y can change, depends on your data
        df.plot(kind='line',x='Iteration',y='Decrease Rate (%)',ax=ax)
        
        # Save current worksheet plot as .png file
        plt.savefig("Sheet"+str(i+1), dpi=None, facecolor='w', edgecolor='w',
                    orientation='portrait', papertype=None, format=None,
                    transparent=False, bbox_inches=None, pad_inches=0.1)
                    
        # Save current worksheet plos as .eps file            
        plt.savefig("Sheet"+str(i+1)+".eps", format='eps')
        
        # Show the plot
        plt.show()       

if __name__ == "__main__":
    main()