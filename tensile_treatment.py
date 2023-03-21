import datetime

import numpy as np
import pandas as pd
import xlwings as xw
from scipy import interpolate
from scipy.optimize import curve_fit
from sklearn.metrics import r2_score
from sklearn.metrics import mean_squared_error as rmse
import matplotlib.pyplot as plt


####################################
# METHODS
####################################
# sugar sintax to write on Excel
def wr(data, row, col):
    ws.range((row, col)).color = (220, 230, 241)
    ws.range((row, col)).value = data


def wr_col(data, first_row, col):
    wr(data.values.reshape((-1, 1)), first_row, col)


# Hyperbolic A model
def HA_fit(ε, a, b):
    return ε / (a + b * ε)


# Hyperbolic B model
def HB_fit(ε, a, b, α):
    return ε / (a + 2 * b * ε) + \
        np.exp(-α * pow(ε - ε_max, 2)) / (2 * b)


# polynomial model
def p6_fit(ε, x6, x5, x4, x3, x2, x1):
    return x6 * (ε ** 6) + \
        x5 * (ε ** 5) + \
        x4 * (ε ** 4) + \
        x3 * (ε ** 3) + \
        x2 * (ε ** 2) + \
        x1 * ε


# find the closest element in a df column
def find_neighbours(value, dataframe, colname):
    exactmatch = dataframe[dataframe[colname] == value]
    if not exactmatch.empty:
        return exactmatch.index
    else:
        lowerneighbour_ind = dataframe[dataframe[colname] < value][colname].idxmax()
        # upperneighbour_ind = dataframe[dataframe[colname] > value][colname].idxmin()
        return lowerneighbour_ind


def set_size(w, h, ax=None):
    """ w, h: width, height in inches """
    if not ax: ax = plt.gca()
    l = ax.figure.subplotpars.left
    r = ax.figure.subplotpars.right
    t = ax.figure.subplotpars.top
    b = ax.figure.subplotpars.bottom
    figw = float(w) / (r - l)
    figh = float(h) / (t - b)
    ax.figure.set_size_inches(figw, figh)


####################################
# PLOTTING CONFIG
####################################

# initialize plot
fig, ax = plt.subplots()

# make fonts editable when saving as pdf
plt.rcParams['pdf.fonttype'] = 42

plt.rc('font', **{'family': 'serif', 'serif': ['Times New Roman']})

# plt.rcParams.update(font)

# plt.title('GG1 experimental data', fontdict=font)
ax.tick_params(direction='in', which='both')
plt.xlabel('Strain, ε [%]')
plt.ylabel('Load per unit width, T [kN/m]')
set_size(2.527, 1.834)
####################################
# DATA TREATMENT
####################################

# select geogrid
geogrid_list = ['sample_experimental_results']
geogrid_treat = True
if geogrid_treat:
    for geogrid in geogrid_list:

        # data file
        RawDataPath = 'D:\\OneDrive\\Documents\\Repos\\' \
                      'Tensile_treatment\\ExperimentalResults\\'
        RawDataName = 'sample_experimental_results' + ".csv"

        # results file
        ResultPath = 'D:\\OneDrive\\Documents\\Repos\\' \
                     'Tensile_Treatment\\TreatedResults\\'
        ResultName = 'sample_treated_results' + ".xlsx"

        # open results file
        wb = xw.Book(ResultPath + ResultName)

        # activate the desired sheet
        ws = wb.sheets['Treated']

        # read the data
        df_raw = pd.read_csv(RawDataPath + RawDataName)

        # get the number of specimens from the raw data file
        n_spc = int(df_raw.shape[1] / 2)

        # write the metadata
        ws.range((2, 4)).value = RawDataPath + RawDataName
        ws.range((3, 4)).value = ResultPath + ResultName
        now = datetime.datetime.now()
        ws.range((4, 4)).value = now.strftime("%Y-%m-%d %H:%M:%S")

        # number of points for interpolation results
        N_POINTS = 40

        # initialize main dataframe
        df_itp = pd.DataFrame()
        σ_f_list = []

        # interpolate every specimen ////////////////////////////////////////////////////
        '''
        To match the following code, the tensile test data is in the following csv format:
        strain 1     stress 1        strain 2       stress 2
        x           x
        x           x
        x           x
        x           x
        
        where the column header defines the pair strain/stress.
        '''

        for spc in range(n_spc):
            # select ε and σ columns
            ε_col = "strain " + str(spc + 1)
            σ_col = "stress " + str(spc + 1)

            # create df without invalid items
            df = df_raw.loc[:, [ε_col, σ_col]][
                ~df_raw.loc[:, [ε_col, σ_col]].isin([np.nan, np.inf, -np.inf]).any(axis=1)]

            # sort the curve by ε
            df = df.sort_values(by=["strain " + str(spc + 1)])
            df.reset_index(drop=True, inplace=True)

            # only the positive data until max σ value
            σ_max = df.iloc[:, 1].max()
            σ_max_idx = df.iloc[:, 1].idxmax()
            ε_min_idx = df[ε_col].abs().idxmin()
            df = df.iloc[ε_min_idx:σ_max_idx, :]

            df.reset_index(drop=True, inplace=True)

            # making sure the ε is strictly increasing
            dx = np.diff(df[ε_col])
            if np.any(dx <= 0):
                for idx, diff in np.ndenumerate(dx):
                    if diff <= 0:
                        df.loc[idx[0], :] = np.NaN
            df.dropna(inplace=True)
            df.reset_index(drop=True, inplace=True)

            # move plot to the (0,0) to better predict the initial stiffness
            first_row_values = df.iloc[[0]].values[0]
            df = df.apply(lambda row: row - first_row_values, axis=1)

            '''
            The next lines do the interpolation:
            σ_f is the function that allows interpolation at any point within the
            domain (above defined as [ε_min_idx:σ_max_idx]). Outside this domain, interpolation
            isn't useful
            '''

            σ_f = interpolate.PchipInterpolator(df[ε_col], df[σ_col], extrapolate=True)
            σ_f_list.append(σ_f)

            # create the ε domain of specimen i
            ε_interp = pd.DataFrame(np.linspace(df[ε_col].head(1), df[ε_col].tail(1), num=N_POINTS),
                                    columns=['ε ' + str(spc + 1)])

            # ε_interp.loc[0.11] = 0.01
            # ε_interp.loc[0.12] = 0.1
            # ε_interp = ε_interp.sort_index().reset_index(drop=True)

            # create the σ image of specimen i
            σ_interp = pd.DataFrame(σ_f(ε_interp),
                                    columns=['σ ' + str(spc + 1)])

            # save on the main interpolated data frame
            df_itp = pd.concat([df_itp, ε_interp], axis=1)
            df_itp = pd.concat([df_itp, σ_interp], axis=1)

            # print the goodness of interpolation just for fun
            wr(r2_score(df[σ_col], σ_f(df[ε_col])), 34, 3 + 2 * spc)

            print_specimen = True
            if print_specimen:
                plt.plot(ε_interp,
                         σ_interp,
                         '.',
                         alpha=0.25,
                         # color='gray',
                         label='GG1' + ' spc ' + str(spc + 1))
        # ///////////////////////////////////////////////////////////////////////////////

        # find median ε_max
        arr = np.arange(len(df_itp.columns)) % 2
        ε_med = float(df_itp.iloc[:, arr == 0].tail(1).mean(axis=1))

        # create median ε
        df_itp_med_2 = pd.DataFrame()
        ε_itp_med = pd.DataFrame(np.linspace(0, ε_med, num=N_POINTS), columns=['ε median'])

        ε_itp_med.loc[0.11] = 0.01
        ε_itp_med.loc[0.12] = 0.1
        ε_itp_med = ε_itp_med.sort_index().reset_index(drop=True)

        # evaluate median σ
        for i in range(len(σ_f_list)):
            σ_itp_med = pd.DataFrame(σ_f_list[i](ε_itp_med), columns=['σ median ' + str(i + 1)])
            df_itp_med_2 = pd.concat([df_itp_med_2, σ_itp_med], axis=1)

        # save the median to the main data frame
        df_itp_med = pd.DataFrame()
        df_itp_med['med ε'] = ε_itp_med
        df_itp_med['med σ'] = df_itp_med_2.iloc[:, :].mean(axis=1)

        # drop the median values after the maximum
        σ_med_max_idx = df_itp_med.iloc[:, 1].idxmax()
        df_itp_med = df_itp_med.iloc[:σ_med_max_idx, :]
        df_itp_med.dropna(inplace=True)

        df_itp = pd.concat([df_itp, df_itp_med], axis=1)

        # create median function
        σ_f_med = interpolate.PchipInterpolator(df_itp_med['med ε'],
                                                df_itp_med['med σ'], extrapolate=True)

        HA_coef, HA_pcov = curve_fit(HA_fit, df_itp_med['med ε'],
                                     df_itp_med['med σ']
                                     )

        HA_perr = np.sqrt(np.diag(HA_pcov))

        ε_max = float(df_itp_med['med ε'].max())
        ε_max_idx = df_itp_med['med ε'].idxmax()
        σ_max = float(df_itp_med['med σ'].max())
        σ_max_idx = df_itp_med['med σ'].idxmax()
        σ_max_HAfit = HA_fit(ε_max, *HA_coef)
        # write stuff in Excel /////////////////////////////////////////////////////////
        # first clean the space and define starting column and row
        write_excel = True
        if write_excel:
            ws.range('B39:T1000').value = None
            SPC_COL = 2
            SPC_ROW = 39
            HA_COL = 19
            HA_ROW = 39

            # write hyperbolic model params
            # the coefficients and covariance
            wr_col(pd.Series(HA_coef.T), 8, 15)
            wr_col(pd.Series(HA_perr.T), 8, 16)
            # the R2 of the median
            wr(r2_score(df_itp_med['med σ'], HA_fit(df_itp_med['med ε'], *HA_coef)), 15, 15)
            wr(rmse(df_itp_med['med σ'], HA_fit(df_itp_med['med ε'], *HA_coef), squared=False), 15, 16)

            # the hyperbolic results of ε and σ
            # check for what strain the hyperbolic model gives the maximum stress
            ε_max_ha = -HA_coef[0] * σ_max / (HA_coef[1] * σ_max - 1)
            ε_ha = pd.DataFrame(np.linspace(0, ε_max_ha, num=N_POINTS))
            # wr_col(ε_ha, SPC_ROW, HA_COL)
            wr_col(df_itp_med['med ε'], SPC_ROW, HA_COL)
            # wr_col(HA_fit(ε_ha, *HA_coef), SPC_ROW, HA_COL + 1)
            wr_col(HA_fit(df_itp_med['med ε'], *HA_coef), SPC_ROW, HA_COL + 1)

            # write results for every specimen
            avg_r2_median = []
            for spc in range(n_spc):
                ε_col = df_itp['ε ' + str(spc + 1)]
                σ_col = df_itp['σ ' + str(spc + 1)]
                wr_col(ε_col, SPC_ROW, SPC_COL + 2 * spc)
                wr_col(σ_col, SPC_ROW, SPC_COL + 2 * spc + 1)

                # R² between the HA model and the specimens
                wr(pd.Series(r2_score(σ_col,
                                      HA_fit(ε_col,
                                             *HA_coef)))[0], 16 + spc, 15)

                # RMSE between the HA model and the specimens
                wr(pd.Series(rmse(σ_col,
                                  HA_fit(ε_col,
                                         *HA_coef), squared=False))[0], 16 + spc, 16)

                # R² between the median and the specimens
                avg_r2_median.append(r2_score(σ_col, σ_f_med(ε_col)))

                if spc == n_spc - 1:
                    wr_col(df_itp_med['med ε'], SPC_ROW, SPC_COL + 2 * i + 2)
                    wr_col(df_itp_med['med σ'], SPC_ROW, SPC_COL + 2 * i + 3)

            # write the average R² between the median and specimens
            wr(np.mean(avg_r2_median), 34, 13)
        # ///////////////////////////////////////////////////////////////////////////////

        ####################################
        # PLOTS
        ####################################

        print_global_fit = False
        if print_global_fit:

            match geogrid:
                case 'GG1':
                    marker = 'o'
                    geogrid_name = 'GG1'
                    xpos = 2
                case 'GG2':
                    marker = '^'
                    geogrid_name = 'GG2'
                    xpos = 2
                case 'GG3':
                    marker = 's'
                    geogrid_name = 'GG3'
                    xpos = 4

            plt.plot(df_itp['med ε'],
                     df_itp['med σ'],
                     marker,
                     markerfacecolor='none',
                     markeredgewidth=0.25,
                     # color='black',
                     label=geogrid_name + ' mean')

            # ax.plot(ε_max, σ_max, 'o', color='black')
            # ax.text(ε_max - 0, σ_max - 4, '$T_{max}$: %.2f kN/m' % σ_max)
            # ax.text(ε_max - 0, σ_max - 10, '$T_{max}$: %.2f kN/m (HA)' % σ_max_HAfit)
            # ax.text(ε_max - 0, σ_max - 7, '$ε_{max}$: %.2f %%' % ε_max)

print_fit = False
if print_fit:
    match geogrid:
        case 'SS20':
            marker = 'o'
            geogrid_name = 'GG1'
        case 'SS40':
            marker = '^'
            geogrid_name = 'GG2'
        case _:
            marker = 's'
            geogrid_name = 'GG3'

    plt.plot(df_itp['med ε'],
             df_itp['med σ'],
             marker=marker,
             markerfacecolor='none',
             color='black',
             label=geogrid_name + ' mean')

    # Add dot and corresponding text
    ax.plot(ε_max, σ_max, 'o', color='black')
    ax.text(ε_max - 0, σ_max - 4, '$T_{max}$: %.2f kN/m' % σ_max)
    ax.text(ε_max - 0, σ_max - 10, '$T_{max}$: %.2f kN/m (HA)' % σ_max_HAfit)
    ax.text(ε_max - 0, σ_max - 7, '$ε_{max}$: %.2f %%' % ε_max)


plt.legend(loc="lower right", handlelength=1, frameon=False)
plt.show()
