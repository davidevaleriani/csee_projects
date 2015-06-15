import numpy as np
import pandas as pd
from sklearn.metrics import mean_squared_error as mse

# get the linear interpolant
def get_linear_interpolant(x0, x1, y0, y1):
    # x-intercept | (x1 y0-x0 y1)/(y0-y1)
    # y-intercept | (x0 y1-x1 y0)/(x0-x1)
    # slope | (y0-y1)/(x0-x1)
    bias = (x0 * y1 - x1 * y0) / (x0 - x1)
    theta = (y0 - y1) / (x0 - x1)
    # print bias, theta, "bias, theta"
    return bias, theta

# find the area under a linear function
def calculate_linear_integral(a, b, bias, theta):
    def F(x):
        # No need for C, finite integral
        return (x * x * theta) / 2.0 + x * bias
    return F(b) - F(a)


def calculate_area_under_piecewise_linear(num_samples, mses):
    # # pad with zeros
    # num_samples = [0.0] + num_samples
    # # Start with a bad assumotion
    # mses = [1.0] + mses
    # convert to numpy floats, just in case
    num_samples = np.array(num_samples, dtype=float)
    mse0 = np.array(mses, dtype=float)
    # normalise
    num_samples /= num_samples.max()
    # integrals can be decomposed linearly
    total_area_under_pieces = 0.0
    for i in range(len(num_samples) - 1):
        num_sample0, num_sample1 = num_samples[i], num_samples[i + 1]
        mse0, mse1 = mses[i], mses[i + 1]
        bias, theta = get_linear_interpolant(
            num_sample0, num_sample1, mse0, mse1)
        area_under_pieces = calculate_linear_integral(
            num_sample0, num_sample1, bias, theta)
        total_area_under_pieces += area_under_pieces
    return total_area_under_pieces


def load_y_hats(y_hats_dir):
    basic_name = y_hats_dir + "/sample"
    # TODO decide which samples we want
    n_samples = [1000, 10000, 20000]
    values = []
    for filename in [basic_name + "_" + str(sample) + ".csv" for sample in n_samples]:
        try:
            print("Processing filename %s" % filename)
            df = pd.read_csv(filename)
            values.append(df.values.T[0])
        except Exception as e:
            print('INCORRECT SUBMISSION %s' % filename)
            return None, None
    return n_samples, values


def get_score(submission_dir, labels_filename="data/testing_y.csv"):
    n_samples, y_hats = load_y_hats(submission_dir)
    # Load labels
    df = pd.read_csv(labels_filename)
    y = df.values.T[0]

    if n_samples is None:
        return 10.0

    mses = [mse(y, y_hat) for y_hat in y_hats]
    return calculate_area_under_piecewise_linear(n_samples, mses)