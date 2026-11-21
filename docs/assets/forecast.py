from xll_python import function


@function(name="FORECAST.UNITS")
def forecast(units: int, growth: float) -> int:
    """Project unit sales at the given growth rate."""
    return round(units * (1 + growth))
