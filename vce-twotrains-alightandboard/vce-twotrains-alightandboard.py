# This is a recursive peak-hour platform clearance calculator.
# simtime simulation length in seconds
# w is total width of upstairs VCEs in feet.
# model from https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf
import openpyxl
import numpy as np

filepath = "platform_F_twotrains_twoway.xlsx"
wb = openpyxl.load_workbook(filepath)

sheet = wb.active
wb.save(filepath)
wb.close()

simtime = 600  # how many seconds we want to simulate
width = 15
length = 1100
cf = 0.75  # area correction factor
eff_area = width * length * cf
train1arrivingpax = 1200  # train 1 load
train2arrivingpax = 1200  # train 2 load
train1doors = 48  # single-door equivalents
train2doors = 48  # single-door equivalents
arrtime1 = 0  # time of first train arrival
arrtime2 = 120  # time of second train arrival
w = 550 / 12
ww = (
    1
    / 12
    * np.array(
        [
            [60, 60, 36, 36, 40, 54, 40, 64, 54, 54, 54, 54, 54],
            [0.6, 0.6, 0.6, 0.6, 1.0, 1.0, 1.0, 1.0, 1.0, 1.4, 1.4, 1.4, 1.4],
        ]
    )
)
www = ww[0, :]
print(www)

# basic flow: train egress > platform crowd > VCE egress rate > back to platform crowd
# if t2 >= simtime, only consider one train.


def deboardratefn(
    k, t, t0, u
):  # of people on train 1 without ingress. t0 is the train arrival time. train(x)doors*rate (1 pax/door/s) goes in u.
    if t >= t0 and k > 0:
        return u
    else:
        return 0


def plat_clearance_fn(
    k, a, w, t, t1
):  # t is i, t1 is the train 1 arrival time, t2 is the train 2 arrival time.
    if k > 0:
        if t - t1 < 30:
            return max(
                min(((111 * (a / k) - 162) / ((a / k) ** 2)) * w / 60, 19 * w / 60), 0
            )  # P = (111M - 162)/(M^2) is the upstairs flow eq per ft wide.
        else:
            return max(
                min(((111 * (a / k) - 162) / ((a / k) ** 2)) * w / 60, 19 * w / 60),
                5 * w / 60,
            )  # Zero clearance rate until train 1 arrives, then hold at a minimum until everyone is upstairs
    else:
        return 0


def plat_ingress_fn(
    r, w
):  # function of platform clearance rate (r), which is itself a function of the platform crowd.
    if r > 17 * w / 60:
        return 0
    elif 5 * w / 60 < r < 17 * w / 60:
        return min(1 * w / 60, max(0, 17 * w / 60 - r * 1.2))
    else:
        return 11 * w / 60


def boardratefn(
    r_deboard, r_platingress, t, t0, u
):  # this r is the deboard rate. Gives the rate at which passengers board each train
    if t > t0 and r_deboard == 0:
        return min(
            r_platingress / 2, u / 8
        )  # arbitrary number of 1 pax every 4 seconds. Assumes passengers partition evenly between the 2 trains.
    else:
        return 0


def space_per_pax_fn(k, a):
    if k > 0:
        return a / k
    else:
        return a


# initialize counters
numonplatform = 0
train1pax = train1arrivingpax
train2pax = train2arrivingpax
train1_remaining_arrivals = train2arrivingpax
train2_remaining_arrivals = train2arrivingpax
for i in range(0, simtime):
    train1offrate = deboardratefn(train1_remaining_arrivals, i, arrtime1, train1doors)
    train1pax -= train1offrate
    train1_remaining_arrivals -= train1offrate
    train2offrate = deboardratefn(train2_remaining_arrivals, i, arrtime2, train2doors)
    train2pax -= train2offrate
    train2_remaining_arrivals -= train2offrate
    numonplatform += train1offrate
    numonplatform += train2offrate
    plat_egress_rate = plat_clearance_fn(numonplatform, eff_area, w, i, arrtime1)
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
    sheet.cell(row=i + 4, column=1).value = i
    sheet.cell(row=i + 4, column=2).value = train1pax
    sheet.cell(row=i + 4, column=3).value = train2pax
    sheet.cell(row=i + 4, column=4).value = train1onrate - train1offrate
    sheet.cell(row=i + 4, column=5).value = train2onrate - train2offrate
    sheet.cell(row=i + 4, column=6).value = (
        plat_ingress_rate + train1offrate + train2offrate
    )
    sheet.cell(row=i + 4, column=9).value = numonplatform
    sheet.cell(row=i + 4, column=10).value = instcrowding
    sheet.cell(row=i + 4, column=7).value = (
        -plat_egress_rate - train1onrate - train2onrate
    )
    sheet.cell(row=i + 4, column=8).value = (
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
        sheet.cell(row=i + 4, column=11).value = "A"
    elif 25 < instcrowding <= 35:
        sheet.cell(row=i + 4, column=11).value = "B"
    elif 15 < instcrowding <= 25:
        sheet.cell(row=i + 4, column=11).value = "C"
    elif 10 < instcrowding <= 15:
        sheet.cell(row=i + 4, column=11).value = "D"
    elif 5 < instcrowding <= 10:
        sheet.cell(row=i + 4, column=11).value = "E"
    else:
        sheet.cell(row=i + 4, column=11).value = "F"

    if plat_egress_rate <= w * 5 / 60:
        sheet.cell(row=i + 4, column=12).value = "A"
    elif w * 5 / 60 < plat_egress_rate <= w * 7 / 60:
        sheet.cell(row=i + 4, column=12).value = "B"
    elif w * 7 / 60 < plat_egress_rate <= w * 9.5 / 60:
        sheet.cell(row=i + 4, column=12).value = "C"
    elif w * 9.5 / 60 < plat_egress_rate <= w * 13 / 60:
        sheet.cell(row=i + 4, column=12).value = "D"
    elif w * 13 / 60 < plat_egress_rate <= w * 17 / 60:
        sheet.cell(row=i + 4, column=12).value = "E"
    else:
        sheet.cell(row=i + 4, column=12).value = "F"

sheet.cell(row=2, column=1).value = "Platform width (ft)"
sheet.cell(row=2, column=2).value = width
sheet.cell(row=2, column=3).value = "Platform length (ft)"
sheet.cell(row=2, column=4).value = length
sheet.cell(row=2, column=5).value = "Total VCE width (ft)"
sheet.cell(row=2, column=6).value = w
sheet.cell(row=2, column=7).value = "Effective Area Multiplier"
sheet.cell(row=2, column=8).value = cf
sheet.cell(row=2, column=9).value = "Usable Platform Area (sqft)"
sheet.cell(row=2, column=10).value = eff_area
sheet.cell(row=2, column=11).value = "Train 1 Arriving Passengers"
sheet.cell(row=2, column=12).value = train1arrivingpax
sheet.cell(row=2, column=13).value = "Train 1 Arrival Time"
sheet.cell(row=2, column=14).value = arrtime1
sheet.cell(row=2, column=15).value = "Train 2 Arriving Passengers"
sheet.cell(row=2, column=16).value = train2arrivingpax
sheet.cell(row=2, column=17).value = "Train 2 Arrival Time"
sheet.cell(row=2, column=18).value = arrtime2
sheet.cell(row=2, column=19).value = "Simulation Length (s)"
sheet.cell(row=2, column=20).value = simtime

sheet.cell(row=3, column=1).value = "Time after arrival (s)"
sheet.cell(row=3, column=2).value = "Passengers on Train 1"
sheet.cell(row=3, column=3).value = "Passengers on Train 2"
sheet.cell(row=3, column=4).value = "Train 1 Board Rate (pax/s)"
sheet.cell(row=3, column=5).value = "Train 2 Board Rate (pax/s)"
sheet.cell(row=3, column=6).value = "Platform Ingress Rate (pax/s)"
sheet.cell(row=3, column=9).value = "Passengers on Platform"
sheet.cell(row=3, column=10).value = "Average Space on Platform per Passenger (sqft)"
sheet.cell(row=3, column=7).value = "Platform Egress Rate (pax/s)"
sheet.cell(row=3, column=8).value = "Net Platform Flow Rate"
sheet.cell(row=3, column=11).value = "Platform Crowding LOS"
sheet.cell(row=3, column=12).value = "Egress LOS"

print(
    "LOS F egress rate is "
    + str(w * 19 / 60)
    + " pax/min. Emergency egress time is roughly "
    + str((train1arrivingpax + train2arrivingpax) / (w * 19 / 60))
    + " seconds."
)
wb.save(filepath)
