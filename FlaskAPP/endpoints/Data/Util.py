import xlsxwriter


def circle_fill(page, testinfo, circles):
    # Hard coded rows implement the header while the for loop inserts the data below the header.
    row = 1  # header rows
    col = 0  # header col

    page.write(0, 0, 'Circle #')
    page.write(0, 1, 'symbol')
    page.write(0, 2, 'Start Time')
    page.write(0, 3, 'End Time')
    page.write(0, 4, 'Total Time')

    page.write(2, 10, 'TestID')
    page.write(2, 11, 'PatientID')
    page.write(2, 12, 'Date')
    page.write(2, 13, 'DoctorID')
    page.write(2, 14, 'Test')
    page.write(3, 15, 'Length')

    page.write(3, 10, testinfo.TestID)
    page.write(3, 11, testinfo.PatientID)
    page.write(3, 12, testinfo.DateTaken)
    page.write(3, 13, testinfo.DoctorID)
    page.write(3, 14, testinfo.TestName)
    page.write(3, 15, testinfo.TestLength)

    for item in circles:  # loop circles 5 columns
        page.write(row, col, item.CircleID + 1)  # converts to 1..n format
        page.write(row, col + 1, item.symbol)
        page.write(row, col + 2, item.begin_circle)
        page.write(row, col + 3, item.end_circle)
        page.write(row, col + 4, item.total_time)
        row = row + 1


def pressure_fill(page, pressure):
    # Header data
    col = 0
    row = 1
    page.write(0, 0, 'Circle #')
    page.write(0, 2, 'Pressure Points')  # merge from first point to last longest row (XX) with merge_range('C1:XX')

    # Prints in first column the total number of points (1..n)
    for i in range(pressure[-1].CircleID + 1):
        page.write(row, col, i + 1)
        row += 1

    col = 2  # sets columns to where data will begin
    row = 1
    i = 0  # increment for pressure index
    for item in pressure:
        page.write(row, col, item.Pressure)
        row = item.CircleID + 1  # converts to 1..n format
        try:  # catches IndexError for pressure[i + 1] at end
            # print('Row: {}  ||  {} :CircleID'.format(row, item.CircleID))
            col = 2 if pressure[i + 1].CircleID + 1 > row else col + 1  # col reset if the next point is in next row
        except IndexError:
            break
        i += 1
