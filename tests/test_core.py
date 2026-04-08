import pytest
from calculator_salar.core import calculate_total


def test_calculate_total_valid():
    assert calculate_total(10, 2) == 20


def test_calculate_total_negative_price():
    with pytest.raises(ValueError):
        calculate_total(-1, 2)


def test_calculate_total_negative_quantity():
    with pytest.raises(ValueError):
        calculate_total(10, -5)