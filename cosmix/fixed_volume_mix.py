import numbers
from typing import List, Union

import numpy as np
from pint import Quantity, Unit
from tabulate import tabulate

from cosmix import ureg
from cosmix.format import Format, format_quantity, gsheets_quantity_format


class MixSpecies(object):
    def __init__(
        self,
        species_name: str,
        stock_conc: Union[Quantity, None],
        target_conc: Union[Quantity, None],
        target_volume: Union[Quantity, None],
    ):
        self.species_name: str = species_name
        self.stock_conc: Union[Quantity, None] = stock_conc
        self.target_conc: Union[Quantity, None] = target_conc
        self.target_volume: Union[Quantity, None] = target_volume


class FixedVolumeMix(object):
    # needed for float comparison
    FLOAT_TOLERANCE_EQ = 1e-4

    def __init__(
        self,
        mix_name: str,
        total_target_volume: Union[numbers.Number, Quantity],
        default_volume_unit: Unit = ureg.microlitres,
        default_conc_unit: Unit = ureg.nanomolar,
    ):

        self.mix_name = mix_name

        if isinstance(total_target_volume, numbers.Number):
            total_target_volume = total_target_volume * default_volume_unit

        self.total_target_volume: Quantity = total_target_volume
        self.default_volume_unit: Unit = default_volume_unit
        self.default_conc_unit: Unit = default_conc_unit

        self.species_list: List[MixSpecies] = []
        self._check_computed_volume()

    def resize(
        self,
        new_total_target_volume: Union[numbers.Number, Quantity],
        use_target_volume=True,
    ):

        if use_target_volume:
            volume_resize_from = self.total_target_volume
        else:
            volume_resize_from = self.computed_volume()

        if isinstance(new_total_target_volume, numbers.Number):
            new_total_target_volume = new_total_target_volume * self.default_volume_unit
        for species in self.species_list:
            species.target_volume = (
                new_total_target_volume / volume_resize_from
            ) * species.target_volume
        self.total_target_volume = new_total_target_volume
        return self

    def computed_volume(self):
        computed_volume: Quantity = 0 * self.default_volume_unit
        for species in self.species_list:
            if species.target_volume is not None:
                computed_volume += species.target_volume
        return computed_volume

    def check_target_volume_is_met(self):
        if not np.isclose(
            self.computed_volume(), self.total_target_volume, self.FLOAT_TOLERANCE_EQ
        ):
            raise ValueError(
                f"The mix's actual volume {self.computed_volume()} is not equal to the set target volume {self.total_target_volume} (up to {self.FLOAT_TOLERANCE_EQ})"
            )

    def _check_computed_volume(self):
        if self.computed_volume() > self.total_target_volume and not np.isclose(
            self.computed_volume(), self.total_target_volume, self.FLOAT_TOLERANCE_EQ
        ):
            raise ValueError(
                f"The mix's actual volume {self.computed_volume()} is bigger than the set target volume {self.total_target_volume}"
            )

    def add_species(
        self,
        species_name: str,
        stock_conc: Union[None, numbers.Number, Quantity],
        target_conc: Union[None, numbers.Number, Quantity],
        target_volume: Union[None, numbers.Number, Quantity] = None,
    ):

        if isinstance(stock_conc, numbers.Number):
            stock_conc *= self.default_conc_unit
        if isinstance(target_conc, numbers.Number):
            target_conc *= self.default_conc_unit
        if isinstance(target_volume, numbers.Number):
            target_volume *= self.default_volume_unit

        if target_volume is None:
            if (stock_conc is not None) and (target_conc is not None):
                target_volume = target_conc * self.total_target_volume / stock_conc

        self.species_list.append(
            MixSpecies(species_name, stock_conc, target_conc, target_volume)
        )
        self._check_computed_volume()

    def add_species_relative_to(
        self,
        species_name: str,
        stock_conc: Union[None, numbers.Number, Quantity],
        relative_to_species_name: str,
        excess: numbers.Number,
    ):

        if isinstance(stock_conc, numbers.Number):
            stock_conc *= self.default_conc_unit

        target_conc = None
        for species in self.species_list:
            if species.species_name == relative_to_species_name:
                if species.target_conc is not None:
                    target_conc = excess * species.target_conc
                    target_volume = target_conc * self.total_target_volume / stock_conc
                    break
                else:
                    raise ValueError(
                        f"The species `{relative_to_species_name}` target conc was not set hence we can't compute relative excess"
                    )

        if target_conc is None:
            raise ValueError(
                f"No species `{relative_to_species_name}` was found is the mix hence we can't compute relative excess"
            )

        self.add_species(species_name, stock_conc, target_conc, target_volume)

    def add_species_volume_fraction(
        self,
        species_name: str,
        inverse_fraction: numbers.Number,
    ):
        self.add_species(
            species_name, None, None, self.total_target_volume / inverse_fraction
        )

    def add_species_volume_complete_with(self, species_name):
        if np.isclose(
            self.computed_volume(), self.total_target_volume, self.FLOAT_TOLERANCE_EQ
        ):
            self.add_species(species_name, None, None, 0)
            return
        # Assertion should be true if use has been using the exposed API
        assert self.total_target_volume - self.computed_volume() >= 0
        self.add_species(
            species_name, None, None, self.total_target_volume - self.computed_volume()
        )

    def species_table(
        self, columns_default_unit=False, gsheets_value_and_formats=False
    ):
        table = []
        for species in self.species_list:
            row = []

            for j, q in enumerate(
                [species.stock_conc, species.target_conc, species.target_volume]
            ):
                if q is None:
                    row.append("N/A")
                    continue

                if columns_default_unit:
                    if j in [0, 1]:
                        q = q.to(self.default_conc_unit)
                    else:
                        q = q.to(self.default_volume_unit)
                else:
                    q = q.to_compact()

                val = format_quantity(q)
                if gsheets_value_and_formats:
                    val = q.to_tuple()[0], gsheets_quantity_format(q)
                row.append(val)

            table.append([species.species_name] + row)

        if not columns_default_unit:
            headers = ["Species", "Stock conc", "Target conc", "Volume to move"]
        else:
            headers = [
                "Species",
                "Stock conc ({:~P})".format(self.default_conc_unit),
                "Target conc ({:~P})".format(self.default_conc_unit),
                "Volume to move ({:~P})".format(self.default_volume_unit),
            ]
        return [headers] + table

    def to_ansi_table(self, columns_default_unit=False):
        to_ret = Format.bold + self.mix_name + Format.end + "\n\n"
        to_tabulate = self.species_table(columns_default_unit)

        to_ret += tabulate(to_tabulate[1:], headers=to_tabulate[0]) + "\n"
        to_ret += "\nTotal volume: " + format_quantity(self.computed_volume())
        to_ret += "\nTotal target volume: " + format_quantity(self.total_target_volume)
        return to_ret

    def __str__(self):
        return self.to_ansi_table()
