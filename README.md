# Instructions

1. Install uv: https://docs.astral.sh/uv/getting-started/installation/#installing-uv

2. Ensure you have the following files in the root directory:

- `rosters.xlsx`
- `room_teachers.csv`

2. Run:

```sh
uv run python https://raw.githubusercontent.com/reteps/homeroom/refs/heads/main/gen_all.py
```

3. The output will be in the `out` directory.