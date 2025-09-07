# Instructions

1. Install uv: https://docs.astral.sh/uv/getting-started/installation/#installing-uv

2. Ensure you have the following files in the directory where you run step 3:

- `rosters.xlsx`
- `room_teachers.csv`

3. Run:

```sh
uv run python https://raw.githubusercontent.com/reteps/homeroom/refs/heads/main/gen_all.py
```

4. The output will be in the `out` directory.
