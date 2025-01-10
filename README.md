# Platform Crowd Model

This repo is for modeling platform crowding and alighting and boarding of trains,
specifically at NY Penn Station.
[`vce-two-trains-alight-and-board.py`](./vce-two-trains-alight-and-board.py)
models a single island platform with two tracks.

## Running

To run, we use the Python package manager [`uv`](https://github.com/astral-sh/uv),
which can be installed easily with

```sh
curl -LsSf https://astral.sh/uv/install.sh | sh
```

With `uv` installed, you can then just run the script directly:

```sh
./vce-two-trains-alight-and-board.py
```

In doing so, `uv` will also install dependencies and set up a virtual environment.

## Checking

We also use `ruff` for formatting and linting and `mypy` for type checking.
To run these, which are also checked in CI,
you can run

```sh
uv run ruff format # format
uv run ruff check # lint
uv run mypy . # type check
```
