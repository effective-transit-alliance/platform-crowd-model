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


def main():
    # how many seconds we want to simulate
    simulation_time = 600

    width = 15

    length = 1200

    # area correction factor
    cf = 0.75

    eff_area = width * length * cf

    # train 1 load
    train1_arriving_pax = 1600

    # train 2 load
    train2_arriving_pax = 1600

    # single-door equivalents
    train1_doors = 48

    # single-door equivalents
    train2_doors = 48

    # time of first train arrival
    arr_time1 = 0

    # time of second train arrival
    arr_time2 = 0

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
                [
                    0.6,
                    0.6,
                    0.6,
                    0.6,
                    1.0,
                    1.0,
                    1.0,
                    1.0,
                    1.0,
                    1.4,
                    1.4,
                    1.4,
                    1.4,
                ],
            ]
        )
    )

    www = ww[0, :]

    print("www = ", www)

    # initialize counters
    total_pax_on_platform = 0
    waitingonplatform = 0
    train1_pax = train1_arriving_pax
    train2_pax = train2_arriving_pax
    train1_remaining_arrivals = train2_arriving_pax
    train2_remaining_arrivals = train2_arriving_pax

    file_path = "platform_F_twotrains_twoway_new.xlsx"
    wb = openpyxl.Workbook()

    sheet = wb.active

    rownum = 0

    def make_row(value, description: str):
        nonlocal rownum
        rownum = rownum + 1
        sheet.cell(column=21, row=rownum).value = description
        sheet.cell(column=22, row=rownum).value = value

    make_row(width, "Platform width (ft)")
    make_row(length, "Platform length (ft)")
    make_row(w, "Total VCE width (ft)")
    make_row(cf, "Effective Area Multiplier")
    make_row(eff_area, "Usable Platform Area (sqft)")
    make_row(train1_arriving_pax, "Train 1 Arriving Passengers")
    make_row(arr_time1, "Train 1 Arrival Time")
    make_row(train2_arriving_pax, "Train 2 Arriving Passengers")
    make_row(arr_time2, "Train 2 Arrival Time")
    make_row(simulation_time, "Simulation Length (s)")
    make_row(w * 19 / 60, "LOS F Egress Rate (pax/s)")
    make_row(
        (train1_arriving_pax + train2_arriving_pax) / (w * 19 / 60),
        "Emergency Egress Time (s)",
    )

    del rownum

    class Columns:
        pass

    columns = Columns()

    colnum = 0

    def make_column(field_name, description: str):
        nonlocal colnum
        colnum = colnum + 1
        setattr(columns, field_name, colnum)
        sheet.cell(
            row=1, column=getattr(columns, field_name)
        ).value = description

    make_column("time_after", "Time after arrival (s)")
    make_column("train1_pax", "Passengers on Train 1")
    make_column("train2_pax", "Passengers on Train 2")
    make_column("train1_board_rate", "Train 1 Board Rate (pax/s)")
    make_column("train2_board_rate", "Train 2 Board Rate (pax/s)")
    make_column("train1_arriving_pax", "Platform Ingress Rate (pax/s)")
    make_column("net_flow_rate", "Passengers on Platform")
    make_column("total_pax_on_platform", "Platform Space per Passanger (sqft)")
    make_column("inst_crowding", "Platform Egress Rate (pax/s)")
    make_column("net_pax_flow_rate", "Net Platform Flow Rate")
    make_column("plat_crowd_los", "Platform Crowding LOS")
    make_column("egress_los", "Egress LOS")

    del colnum

    FIRST_DATA_ROW = 2

    for time_after in range(0, simulation_time):
        train1_off_rate = deboardratefn(
            train1_remaining_arrivals, time_after, arr_time1, train1_doors
        )
        train1_pax -= train1_off_rate
        train1_remaining_arrivals -= train1_off_rate
        train2_off_rate = deboardratefn(
            train2_remaining_arrivals, time_after, arr_time2, train2_doors
        )
        train2_pax -= train2_off_rate
        train2_remaining_arrivals -= train2_off_rate
        total_pax_on_platform += train1_off_rate
        total_pax_on_platform += train2_off_rate
        plat_egress_rate = plat_clearance_fn(
            total_pax_on_platform,
            eff_area,
            w,
            waitingonplatform,
            sum(www) * queue_length,
        )
        total_pax_on_platform -= plat_egress_rate
        plat_ingress_rate = plat_ingress_fn(plat_egress_rate, w)
        total_pax_on_platform += plat_ingress_rate
        train1_on_rate = boardratefn(
            train1_off_rate,
            plat_ingress_rate,
            time_after,
            arr_time1,
            train1_doors,
        )
        train2_on_rate = boardratefn(
            train2_off_rate,
            plat_ingress_rate,
            time_after,
            arr_time2,
            train2_doors,
        )

        if train1_pax <= train1_arriving_pax:
            train1_pax += train1_on_rate
            total_pax_on_platform -= train1_on_rate
        else:
            train1_pax += 0
            total_pax_on_platform -= 0

        if train2_pax <= train2_arriving_pax:
            train2_pax += train2_on_rate
            total_pax_on_platform -= train2_on_rate
        else:
            train1_pax += 0
            total_pax_on_platform -= 0

        inst_crowding = space_per_pax_fn(total_pax_on_platform, eff_area)
        if total_pax_on_platform < 0:
            total_pax_on_platform = 0

        waitingonplatform = (
            waitingonplatform
            + train1_off_rate
            + train2_off_rate
            - plat_egress_rate
        )
        if waitingonplatform < 0:
            waitingonplatform = 0
        print(
            "time " + str(time_after),
            "train 1 load " + str(train1_pax),
            "train 2 load " + str(train2_pax),
            train1_off_rate,
            train2_off_rate,
            total_pax_on_platform,
            inst_crowding,
            plat_egress_rate,
        )
        row = time_after + FIRST_DATA_ROW
        sheet.cell(row=row, column=1).value = time_after
        sheet.cell(row=row, column=2).value = train1_pax
        sheet.cell(row=row, column=3).value = train2_pax
        sheet.cell(row=row, column=4).value = train1_on_rate - train1_off_rate
        sheet.cell(row=row, column=5).value = train2_on_rate - train2_off_rate
        sheet.cell(row=row, column=6).value = (
            plat_ingress_rate + train1_off_rate + train2_off_rate
        )
        sheet.cell(row=row, column=9).value = total_pax_on_platform
        sheet.cell(row=row, column=10).value = inst_crowding
        sheet.cell(row=row, column=7).value = (
            -plat_egress_rate - train1_on_rate - train2_on_rate
        )
        net_flow_rate = (
            plat_ingress_rate
            + train1_off_rate
            + train2_off_rate
            - plat_egress_rate
            - train1_on_rate
            - train2_on_rate
        )
        sheet.cell(row=row, column=8).value = net_flow_rate
        egr = plat_egress_rate / width * www
        print(egr, np.sum(egr))
        if inst_crowding > 35:
            sheet.cell(row=row, column=11).value = "A"
        elif 25 < inst_crowding <= 35:
            sheet.cell(row=row, column=11).value = "B"
        elif 15 < inst_crowding <= 25:
            sheet.cell(row=row, column=11).value = "C"
        elif 10 < inst_crowding <= 15:
            sheet.cell(row=row, column=11).value = "D"
        elif 5 < inst_crowding <= 10:
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
        + str((train1_arriving_pax + train2_arriving_pax) / (w * 19 / 60))
        + " seconds."
    )
    wb.save(file_path)
    wb.close()


if __name__ == "__main__":
    main()
