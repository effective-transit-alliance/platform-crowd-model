"""
This is a recursive peak-hour platform clearance calculator.
model from https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf
"""

import dataclasses
import numpy as np
import openpyxl
from openpyxl.cell import Cell
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


def plat_crowd_grade(inst_crowding: float) -> str:
    if inst_crowding > 35:
        return "A"
    elif 25 < inst_crowding <= 35:
        return "B"
    elif 15 < inst_crowding <= 25:
        return "C"
    elif 10 < inst_crowding <= 15:
        return "D"
    elif 5 < inst_crowding <= 10:
        return "E"
    else:
        return "F"


def egress_crowd_grade(w, plat_egress_rate: float) -> str:
    if plat_egress_rate <= w * 5 / 60:
        return "A"
    elif w * 5 / 60 < plat_egress_rate <= w * 7 / 60:
        return "B"
    elif w * 7 / 60 < plat_egress_rate <= w * 9.5 / 60:
        return "C"
    elif w * 9.5 / 60 < plat_egress_rate <= w * 13 / 60:
        return "D"
    elif w * 13 / 60 < plat_egress_rate <= w * 17 / 60:
        return "E"
    else:
        return "F"


@dataclasses.dataclass
class Params:
    # how many seconds we want to simulate
    simulation_time: int

    width: int

    length: int

    # area correction factor
    cf: float

    # train 1 load
    train1_arriving_pax: int

    # train 2 load
    train2_arriving_pax: int

    # single-door equivalents
    train1_doors: int

    # single-door equivalents
    train2_doors: int

    # time of first train arrival
    arr_time1: int

    # time of second train arrival
    arr_time2: int

    # feet of queue in front of each stair that pushes max upstairs flow
    queue_length: int

    # total width of upstairs VCEs in feet.
    w: float

    ww: any


def calc_workbook(params: Params) -> openpyxl.Workbook:
    eff_area = params.width * params.length * params.cf

    www = params.ww[0, :]

    print("www = ", www)

    # initialize counters
    total_pax_on_platform = 0
    waitingonplatform = 0
    train1_pax = params.train1_arriving_pax
    train2_pax = params.train2_arriving_pax
    train1_remaining_arrivals = params.train2_arriving_pax
    train2_remaining_arrivals = params.train2_arriving_pax

    wb = openpyxl.Workbook()

    sheet = wb.active

    rownum = 0

    def make_row(value, description: str):
        nonlocal rownum
        rownum = rownum + 1
        sheet.cell(column=21, row=rownum).value = description
        sheet.cell(column=22, row=rownum).value = value

    make_row(params.width, "Platform width (ft)")
    make_row(params.length, "Platform length (ft)")
    make_row(params.w, "Total VCE width (ft)")
    make_row(params.cf, "Effective Area Multiplier")
    make_row(eff_area, "Usable Platform Area (sqft)")
    make_row(params.train1_arriving_pax, "Train 1 Arriving Passengers")
    make_row(params.arr_time1, "Train 1 Arrival Time")
    make_row(params.train2_arriving_pax, "Train 2 Arriving Passengers")
    make_row(params.arr_time2, "Train 2 Arrival Time")
    make_row(params.simulation_time, "Simulation Length (s)")
    make_row(params.w * 19 / 60, "LOS F Egress Rate (pax/s)")
    make_row(
        (params.train1_arriving_pax + params.train2_arriving_pax)
        / (params.w * 19 / 60),
        "Emergency Egress Time (s)",
    )

    del rownum

    class Columns:
        pass

    columns = Columns()

    colnum = 0

    def make_column_num(description: str):
        nonlocal colnum
        colnum = colnum + 1
        sheet.cell(row=1, column=colnum).value = description
        return colnum

    columns.time_after = make_column_num("Time after arrival (s)")
    columns.train1_pax = make_column_num("Passengers on Train 1")
    columns.train2_pax = make_column_num("Passengers on Train 2")
    columns.train1_board_rate = make_column_num("Train 1 Board Rate (pax/s)")
    columns.train2_board_rate = make_column_num("Train 2 Board Rate (pax/s)")
    columns.plat_ingress_rate = make_column_num(
        "Platform Ingress Rate (pax/s)"
    )
    columns.plat_egress_rate = make_column_num("Platform Egress Rate (pax/s)")
    columns.total_pax_on_platform = make_column_num("Passengers on Platform")
    columns.inst_crowding = make_column_num(
        "Platform Space per Passanger (sqft)"
    )
    columns.net_pax_flow_rate = make_column_num("Net Platform Flow Rate")
    columns.plat_crowd_los = make_column_num("Platform Crowding LOS")
    columns.egress_los = make_column_num("Egress LOS")

    del colnum

    FIRST_DATA_ROW = 2

    for time_after in range(0, params.simulation_time):
        train1_off_rate = deboardratefn(
            train1_remaining_arrivals,
            time_after,
            params.arr_time1,
            params.train1_doors,
        )
        train1_pax -= train1_off_rate
        train1_remaining_arrivals -= train1_off_rate
        train2_off_rate = deboardratefn(
            train2_remaining_arrivals,
            time_after,
            params.arr_time2,
            params.train2_doors,
        )
        train2_pax -= train2_off_rate
        train2_remaining_arrivals -= train2_off_rate
        total_pax_on_platform += train1_off_rate
        total_pax_on_platform += train2_off_rate
        plat_egress_rate = plat_clearance_fn(
            total_pax_on_platform,
            eff_area,
            params.w,
            waitingonplatform,
            sum(www) * params.queue_length,
        )
        total_pax_on_platform -= plat_egress_rate
        plat_ingress_rate = plat_ingress_fn(plat_egress_rate, params.w)
        total_pax_on_platform += plat_ingress_rate
        train1_on_rate = boardratefn(
            train1_off_rate,
            plat_ingress_rate,
            time_after,
            params.arr_time1,
            params.train1_doors,
        )
        train2_on_rate = boardratefn(
            train2_off_rate,
            plat_ingress_rate,
            time_after,
            params.arr_time2,
            params.train2_doors,
        )

        if train1_pax <= plat_ingress_rate:
            train1_pax += train1_on_rate
            total_pax_on_platform -= train1_on_rate
        else:
            train1_pax += 0
            total_pax_on_platform -= 0

        if train2_pax <= params.train2_arriving_pax:
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

        def get_cell(column: int) -> Cell:
            return sheet.cell(row=row, column=column)

        get_cell(columns.time_after).value = time_after
        get_cell(columns.train1_pax).value = train1_pax
        get_cell(columns.train2_pax).value = train2_pax
        get_cell(columns.train1_board_rate).value = (
            train1_on_rate - train1_off_rate
        )
        get_cell(columns.train2_board_rate).value = (
            train2_on_rate - train2_off_rate
        )
        get_cell(columns.plat_ingress_rate).value = (
            plat_ingress_rate + train1_off_rate + train2_off_rate
        )
        get_cell(columns.total_pax_on_platform).value = total_pax_on_platform
        get_cell(columns.inst_crowding).value = inst_crowding
        get_cell(columns.plat_egress_rate).value = (
            -plat_egress_rate - train1_on_rate - train2_on_rate
        )
        net_pax_flow_rate = (
            plat_ingress_rate
            + train1_off_rate
            + train2_off_rate
            - plat_egress_rate
            - train1_on_rate
            - train2_on_rate
        )
        get_cell(columns.net_pax_flow_rate).value = net_pax_flow_rate
        egr = plat_egress_rate / params.width * www
        print(egr, np.sum(egr))

        get_cell(columns.plat_crowd_los).value = plat_crowd_grade(
            inst_crowding
        )

        get_cell(columns.egress_los).value = egress_crowd_grade(
            params.w, plat_egress_rate
        )

    def make_chart(title, min_col):
        chart = openpyxl.chart.ScatterChart()
        chart.title = title
        chart.style = 13
        chart.x_axis.title = "Size"
        chart.y_axis.title = "Percentage"

        max_row = params.simulation_time + FIRST_DATA_ROW - 1
        xvalues = Reference(
            sheet, min_col=1, min_row=FIRST_DATA_ROW, max_row=max_row
        )
        values = Reference(
            sheet, min_col=min_col, min_row=FIRST_DATA_ROW, max_row=max_row
        )
        series = Series(values, xvalues, title_from_data=True)
        chart.series.append(series)
        return chart

    sheet.add_chart(
        make_chart("Net Platform Flow Rate", columns.net_pax_flow_rate), "M5"
    )
    sheet.add_chart(
        make_chart("Passengers on Platform", columns.total_pax_on_platform),
        "M25",
    )
    sheet.add_chart(
        make_chart("Space per Passenger", columns.inst_crowding), "M45"
    )

    print(
        "LOS F egress rate is "
        + str(params.w * 19 / 60)
        + " pax/second. Emergency egress time is roughly "
        + str(
            (plat_ingress_rate + params.train2_arriving_pax)
            / (params.w * 19 / 60)
        )
        + " seconds."
    )
    return wb


def main():
    wb = calc_workbook(
        params=Params(
            simulation_time=600,
            width=15,
            length=1200,
            cf=0.75,
            train1_arriving_pax=1600,
            train2_arriving_pax=1600,
            train1_doors=48,
            train2_doors=48,
            arr_time1=0,
            arr_time2=0,
            queue_length=20,
            w=550 / 12,
            ww=(
                1
                / 12
                * np.transpose(
                    np.array(
                        [
                            # Width weight pairs
                            [60, 0.6],
                            [60, 0.6],
                            [36, 0.6],
                            [36, 0.6],
                            [40, 1.0],
                            [54, 1.0],
                            [40, 1.0],
                            [64, 1.0],
                            [54, 1.0],
                            [54, 1.4],
                            [54, 1.4],
                            [54, 1.4],
                            [54, 1.4],
                        ]
                    )
                )
            ),
        ),
    )
    wb.save("platform_F_twotrains_twoway_1600_120s.xlsx")
    wb.close()


if __name__ == "__main__":
    main()
