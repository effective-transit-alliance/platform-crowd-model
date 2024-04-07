"""
This is a recursive peak-hour platform clearance calculator.
model from https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf
"""

import numpy as np
import openpyxl
from openpyxl.chart import Reference, Series


# basic flow: train egress > platform crowd > VCE egress rate > back to
# platform crowd

# if t2 >= simulation_time, only consider one train.

# keep high VCE egress rate if queues at stairs are long


def deboardratefn(k, t, t0, u):
    """
    :param k: number of people on given train without ingress
    :param t: time pass counter (s)
    :param t0: train arrival time
    :param u: train(x)doors*rate (1 pax/door/s)
    :return: egress rate from train to platform across all doors (pax/s)
    """
    if t >= t0 and k > 0:
        return u
    else:
        return 0


def plat_clearance_fn(k, a, w, karr, qmax):
    """
    :param k: number of people on the platform
    :param a: usable platform area
    :param w: total width of vertical circulation elements
    :param: karr: number of people waiting to get onto a stairwell
    :param: qmax: number of people that can fit around stair thresholds
    :return: platform egress rate on stairs
    """

    if karr <= qmax:
        # P = (111M - 162)/(M^2) is the upstairs flow eq per ft wide.
        return max(
            0,
            min(
                (111 * a / (max(1, k)) - 162) / ((a / (max(1, k))) ** 2),
                19 * w / 60,
            ),
        )

    elif karr > qmax:
        # Waiting volume over threshold maintains miminum 7 pax/ft/min,
        # which is LOS B/C boundary.
        return max(
            7 * w / 60,
            min(
                (111 * a / (max(1, k)) - 162) / ((a / (max(1, k))) ** 2),
                19 * w / 60,
            ),
        )
    else:
        return 0


def plat_ingress_fn(r, w):
    """
    :param r: platform egress rate, itself a function of the platform
    crowd

    :param w: total width of vertical circulation elements

    :return: platform ingress rate across stairs (pax/s)
    """
    if r > 17 * w / 60:
        return 0
    elif 5 * w / 60 < r < 17 * w / 60:
        return min(1 * w / 60, max(0, 17 * w / 60 - r * 1.2))
    else:
        return 11 * w / 60


def boardratefn(r_deboard, r_platingress, t, t0, u):
    """
    :param r_deboard: train deboard rate
    :param r_platingress: platform ingress rate
    :param t: time in seconds, pass counter
    :param t0: train arrival time
    :return: train ingress rate across all doors (pax/s)
    """
    if t > t0 and r_deboard == 0:
        # arbitrary number of 1 pax every 4 seconds. Assumes passengers
        # partition evenly between the 2 trains.
        return min(r_platingress / 2, u / 8)
    else:
        return 0


def space_per_pax_fn(k, a):
    """
    :param k: people on platform (pax)
    :param a: usable platform area (ft^2)
    :return: space per passenger (ft^2/pax)
    """
    if k > 0:
        return a / k
    else:
        return a


# how many seconds we want to simulate
simulation_time = 600

width = 15

length = 1200

# area correction factor
cf = 0.75

eff_area = width * length * cf

# train 1 load
train1arrivingpax = 1600

# train 2 load
train2arrivingpax = 1600

# single-door equivalents
train1doors = 48

# single-door equivalents
train2doors = 48

# time of first train arrival
arrtime1 = 0

# time of second train arrival
arrtime2 = 0

# feet of queue in front of each stair that pushes max upstairs flow
queue_length = 20

# total width of upstairs VCEs in feet.
w = 550 / 12

ww = (
    1
    / 12
    * np.array(
        [
            # Widths
            [60, 60, 36, 36, 40, 54, 40, 64, 54, 54, 54, 54, 54],
            # Weights
            [0.6, 0.6, 0.6, 0.6, 1.0, 1.0, 1.0, 1.0, 1.0, 1.4, 1.4, 1.4, 1.4],
        ]
    )
)

www = ww[0, :]

print("www = ", www)

# initialize counters
numonplatform = 0
waitingonplatform = 0
train1pax = train1arrivingpax
train2pax = train2arrivingpax
train1_remaining_arrivals = train2arrivingpax
train2_remaining_arrivals = train2arrivingpax

filepath = "platform_F_twotrains_twoway_new.xlsx"
wb = openpyxl.Workbook()

sheet = wb.active

sheet.cell(row=1, column=1).value = "Time after arrival (s)"
sheet.cell(row=1, column=2).value = "Passengers on Train 1"
sheet.cell(row=1, column=3).value = "Passengers on Train 2"
sheet.cell(row=1, column=4).value = "Train 1 Board Rate (pax/s)"
sheet.cell(row=1, column=5).value = "Train 2 Board Rate (pax/s)"
sheet.cell(row=1, column=6).value = "Platform Ingress Rate (pax/s)"
sheet.cell(row=1, column=9).value = "Passengers on Platform"
sheet.cell(row=1, column=10).value = "Platform Space per Passanger (sqft)"
sheet.cell(row=1, column=7).value = "Platform Egress Rate (pax/s)"
sheet.cell(row=1, column=8).value = "Net Platform Flow Rate"
sheet.cell(row=1, column=11).value = "Platform Crowding LOS"
sheet.cell(row=1, column=12).value = "Egress LOS"

sheet.cell(column=21, row=1).value = "Platform width (ft)"
sheet.cell(column=22, row=1).value = width
sheet.cell(column=21, row=2).value = "Platform length (ft)"
sheet.cell(column=22, row=2).value = length
sheet.cell(column=21, row=3).value = "Total VCE width (ft)"
sheet.cell(column=22, row=3).value = w
sheet.cell(column=21, row=4).value = "Effective Area Multiplier"
sheet.cell(column=22, row=4).value = cf
sheet.cell(column=21, row=5).value = "Usable Platform Area (sqft)"
sheet.cell(column=22, row=5).value = eff_area
sheet.cell(column=21, row=6).value = "Train 1 Arriving Passengers"
sheet.cell(column=22, row=6).value = train1arrivingpax
sheet.cell(column=21, row=7).value = "Train 1 Arrival Time"
sheet.cell(column=22, row=7).value = arrtime1
sheet.cell(column=21, row=8).value = "Train 2 Arriving Passengers"
sheet.cell(column=22, row=8).value = train2arrivingpax
sheet.cell(column=21, row=9).value = "Train 2 Arrival Time"
sheet.cell(column=22, row=9).value = arrtime2
sheet.cell(column=21, row=10).value = "Simulation Length (s)"
sheet.cell(column=22, row=10).value = simulation_time
sheet.cell(column=21, row=11).value = "LOS F Egress Rate (pax/s)"
sheet.cell(column=22, row=11).value = w * 19 / 60
sheet.cell(column=21, row=12).value = "Emergency Egress Time (s)"
sheet.cell(column=22, row=12).value = (
    train1arrivingpax + train2arrivingpax
) / (w * 19 / 60)


FIRST_DATA_ROW = 2

for i in range(0, simulation_time):
    train1offrate = deboardratefn(
        train1_remaining_arrivals, i, arrtime1, train1doors
    )
    train1pax -= train1offrate
    train1_remaining_arrivals -= train1offrate
    train2offrate = deboardratefn(
        train2_remaining_arrivals, i, arrtime2, train2doors
    )
    train2pax -= train2offrate
    train2_remaining_arrivals -= train2offrate
    numonplatform += train1offrate
    numonplatform += train2offrate
    plat_egress_rate = plat_clearance_fn(
        numonplatform, eff_area, w, waitingonplatform, sum(www) * queue_length
    )
    numonplatform -= plat_egress_rate
    plat_ingress_rate = plat_ingress_fn(plat_egress_rate, w)
    numonplatform += plat_ingress_rate
    train1onrate = boardratefn(
        train1offrate, plat_ingress_rate, i, arrtime1, train1doors
    )
    train2onrate = boardratefn(
        train2offrate, plat_ingress_rate, i, arrtime2, train2doors
    )

    if train1pax <= train1arrivingpax:
        train1pax += train1onrate
        numonplatform -= train1onrate
    else:
        train1pax += 0
        numonplatform -= 0

    if train2pax <= train2arrivingpax:
        train2pax += train2onrate
        numonplatform -= train2onrate
    else:
        train1pax += 0
        numonplatform -= 0

    instcrowding = space_per_pax_fn(numonplatform, eff_area)
    if numonplatform < 0:
        numonplatform = 0

    waitingonplatform = (
        waitingonplatform + train1offrate + train2offrate - plat_egress_rate
    )
    if waitingonplatform < 0:
        waitingonplatform = 0
    print(
        "time " + str(i),
        "train 1 load " + str(train1pax),
        "train 2 load " + str(train2pax),
        train1offrate,
        train2offrate,
        numonplatform,
        instcrowding,
        plat_egress_rate,
    )
    row = i + FIRST_DATA_ROW
    sheet.cell(row=row, column=1).value = i
    sheet.cell(row=row, column=2).value = train1pax
    sheet.cell(row=row, column=3).value = train2pax
    sheet.cell(row=row, column=4).value = train1onrate - train1offrate
    sheet.cell(row=row, column=5).value = train2onrate - train2offrate
    sheet.cell(row=row, column=6).value = (
        plat_ingress_rate + train1offrate + train2offrate
    )
    sheet.cell(row=row, column=9).value = numonplatform
    sheet.cell(row=row, column=10).value = instcrowding
    sheet.cell(row=row, column=7).value = (
        -plat_egress_rate - train1onrate - train2onrate
    )
    sheet.cell(row=row, column=8).value = (
        plat_ingress_rate
        + train1offrate
        + train2offrate
        - plat_egress_rate
        - train1onrate
        - train2onrate
    )
    egr = plat_egress_rate / width * www
    print(egr, np.sum(egr))
    if instcrowding > 35:
        sheet.cell(row=row, column=11).value = "A"
    elif 25 < instcrowding <= 35:
        sheet.cell(row=row, column=11).value = "B"
    elif 15 < instcrowding <= 25:
        sheet.cell(row=row, column=11).value = "C"
    elif 10 < instcrowding <= 15:
        sheet.cell(row=row, column=11).value = "D"
    elif 5 < instcrowding <= 10:
        sheet.cell(row=row, column=11).value = "E"
    else:
        sheet.cell(row=row, column=11).value = "F"

    if plat_egress_rate <= w * 5 / 60:
        sheet.cell(row=row, column=12).value = "A"
    elif w * 5 / 60 < plat_egress_rate <= w * 7 / 60:
        sheet.cell(row=row, column=12).value = "B"
    elif w * 7 / 60 < plat_egress_rate <= w * 9.5 / 60:
        sheet.cell(row=row, column=12).value = "C"
    elif w * 9.5 / 60 < plat_egress_rate <= w * 13 / 60:
        sheet.cell(row=row, column=12).value = "D"
    elif w * 13 / 60 < plat_egress_rate <= w * 17 / 60:
        sheet.cell(row=row, column=12).value = "E"
    else:
        sheet.cell(row=row, column=12).value = "F"

def make_chart(title, min_col):
    chart = openpyxl.chart.ScatterChart()
    chart.title = title
    chart.style = 13
    chart.x_axis.title = "Size"
    chart.y_axis.title = "Percentage"

    max_row = simulation_time + FIRST_DATA_ROW - 1
    xvalues = Reference(
        sheet, min_col=1, min_row=FIRST_DATA_ROW, max_row=max_row
    )
    values = Reference(
        sheet, min_col=min_col, min_row=FIRST_DATA_ROW, max_row=max_row
    )
    series = Series(values, xvalues, title_from_data=True)
    chart.series.append(series)
    return chart


sheet.add_chart(make_chart("Net Platform Flow Rate", 8), "M5")
sheet.add_chart(make_chart("Passengers on Platform", 9), "M25")
sheet.add_chart(make_chart("Space per Passenger", 10), "M45")

print(
    "LOS F egress rate is "
    + str(w * 19 / 60)
    + " pax/second. Emergency egress time is roughly "
    + str((train1arrivingpax + train2arrivingpax) / (w * 19 / 60))
    + " seconds."
)
wb.save(filepath)
wb.close()
