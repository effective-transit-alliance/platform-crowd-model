train1arr = 1600
train2arr = 1600
platarea = 15 * 900
cf = 0.75
corrarea = platarea * cf
arrtime1 = 0  # time when train 1 opens doors in seconds
arrtime2 = 120  # time when train 2 opens doors in seconds
simtime = 600
train1space = 85 * 9 * 10  # l x w x cars in sqft
train2space = 85 * 9 * 10  # l x w x cars in sqft
train1doorwidth = 8 * 10
train2doorwidth = 8 * 10  # total width of doors

# initialize counter variables
train1load = train1arr
train2load = train2arr
numonplatform = 0  # consider calling "arrpaxonplatform
w = 500 / 12


def train_alight_fn(
    k, a, t, t0, d
):  # convert to rate per second by dividing by 60
    if t >= t0 and k > 0:
        return max(
            min(
                ((267 * (a / k) - 722) / ((a / k) ** 2)) * d / 60, 25 * d / 60
            ),
            5 * d / 60,
        )  # P = (267M - 722)/(M^2) is the bidirectional flow eq. per ft wide. M is sqft per ped
    else:
        return 0


def plat_clearance_fn(k, a, w):
    if k > 0:
        return max(
            min(
                ((111 * (a / k) - 162) / ((a / k) ** 2)) * w / 60, 19 * w / 60
            ),
            5 * w / 60,
        )  # P = (111M - 162)/(M^2) is the upstairs flow eq per ft wide.
    else:
        return 0


for i in range(1, simtime):
    train1offrate = train_alight_fn(
        train1load, train1space, i, arrtime1, train1doorwidth
    )
    train1load -= train1offrate
    if train1load <= 0:
        train1load = 0
    train2offrate = train_alight_fn(
        train2load, train2space, i, arrtime2, train2doorwidth
    )
    train2load -= train2offrate
    if train2load <= 0:
        train2load = 0
    numonplatform += train1offrate + train2offrate

    if numonplatform > 0:
        plat_space_perpax = corrarea / numonplatform
    else:
        plat_space_perpax = corrarea
    plat_egress_rate = plat_clearance_fn(numonplatform, corrarea, w)
    numonplatform -= plat_egress_rate
    if numonplatform <= 0:
        numonplatform = 0
    print(
        i,
        train1offrate,
        train1load,
        train2offrate,
        train2load,
        numonplatform,
        plat_space_perpax,
        plat_egress_rate,
    )

# https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf
print("LOS F egress rate is " + str(w * 19 / 60) + " pax/sec")
