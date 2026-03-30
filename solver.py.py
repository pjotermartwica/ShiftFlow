"""
Moduł logiki planowania grafiku przy użyciu Google OR-Tools.
Zawiera logikę biznesową niezależną od interfejsu użytkownika.
"""

from ortools.sat.python import cp_model
from typing import Dict, List, Tuple


class ScheduleSolver:
    """
    Klasa do tworzenia grafiku dla zespołu людей.
    Używa Google OR-Tools do optymalizacji przydziału osób do zmian.
    """
    
    def __init__(self, num_people: int = 14, num_days: int = 7, shifts_per_day: int = 3):
        """
        Inicjalizacja solvera.
        
        Args:
            num_people: Liczba osób w zespole (domyślnie 14)
            num_days: Liczba dni w grafiku (domyślnie 7)
            shifts_per_day: Liczba zmian na dzień (domyślnie 3: poranek, popołudnie, noc)
        """
        self.num_people = num_people
        self.num_days = num_days
        self.shifts_per_day = shifts_per_day
        self.model = None
        self.solution = None
        
    def create_schedule(self) -> Dict[str, List[List[str]]]:
        """
        Tworzy grafik dla zespołu.
        
        Returns:
            Dict z kluczami 'schedule' i 'status' zawierający:
            - schedule: Grafik w formacie [dzień][zmiana] = lista osób
            - status: Status rozwiązania ('OPTIMAL', 'FEASIBLE', 'INFEASIBLE')
        """
        self.model = cp_model.CpModel()
        shift_names = ["Poranek", "Popołudnie", "Noc"]
        people_names = [f"Osoba {i+1}" for i in range(self.num_people)]
        
        # Zmienne decyzyjne: (osoba, dzień, zmiana) -> 0/1
        assignments = {}
        for person in range(self.num_people):
            for day in range(self.num_days):
                for shift in range(self.shifts_per_day):
                    assignments[(person, day, shift)] = self.model.NewBoolVar(
                        f"person_{person}_day_{day}_shift_{shift}"
                    )
        
        # Ograniczenie: każda zmiana musi mieć dokładnie 2 osoby (dla 14 osób, 3 zmian)
        for day in range(self.num_days):
            for shift in range(self.shifts_per_day):
                self.model.Add(
                    sum(assignments[(person, day, shift)] 
                        for person in range(self.num_people)) == 2
                )
        
        # Ograniczenie: każda osoba pracuje maksymalnie 3 dni w tygodniu
        max_days_per_person = 3
        for person in range(self.num_people):
            self.model.Add(
                sum(assignments[(person, day, shift)]
                    for day in range(self.num_days)
                    for shift in range(self.shifts_per_day)) <= max_days_per_person * self.shifts_per_day
            )
        
        # Ograniczenie: każda osoba nie pracuje więcej niż 1 zmianę na dzień
        for person in range(self.num_people):
            for day in range(self.num_days):
                self.model.Add(
                    sum(assignments[(person, day, shift)]
                        for shift in range(self.shifts_per_day)) <= 1
                )
        
        # Optymalizacja: maksymalnie równomierny rozkład pracy
        person_shifts = [
            sum(assignments[(person, day, shift)]
                for day in range(self.num_days)
                for shift in range(self.shifts_per_day))
            for person in range(self.num_people)
        ]
        self.model.Minimize(max(person_shifts) - min(person_shifts))
        
        # Rozwiązanie
        solver = cp_model.CpSolver()
        status = solver.Solve(self.model)
        
        # Konwersja wyniku na czytelny format
        schedule = [
            [[] for _ in range(self.shifts_per_day)]
            for _ in range(self.num_days)
        ]
        
        if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            for person in range(self.num_people):
                for day in range(self.num_days):
                    for shift in range(self.shifts_per_day):
                        if solver.Value(assignments[(person, day, shift)]):
                            schedule[day][shift].append(people_names[person])
        
        status_text = "OPTIMAL" if status == cp_model.OPTIMAL else \
                      "FEASIBLE" if status == cp_model.FEASIBLE else \
                      "INFEASIBLE"
        
        return {
            "schedule": schedule,
            "status": status_text,
            "people": people_names,
            "shift_names": shift_names
        }


class ScheduleOptimizer:
    """Prosty optymalizator grafiku dla 14 pracowników na 7 dni z 2 zmianami."""

    def __init__(self,
                 num_people: int = 14,
                 num_days: int = 7,
                 shifts_per_day: int = 2,
                 preferences: Dict[int, List[int]] = None):
        self.num_people = num_people
        self.num_days = num_days
        self.shifts_per_day = shifts_per_day
        self.preferences = preferences or {}  # {pracownik: [preferowane_zmiany]}

        self.model = cp_model.CpModel()
        self.assignments = {}

        # Jednostki zmian
        self.shift_names = ["Poranna", "Popołudniowa"]

        # Zmienna decyzyjna: czy osoba p pracuje w dniu d na zmianie s
        for p in range(self.num_people):
            for d in range(self.num_days):
                for s in range(self.shifts_per_day):
                    self.assignments[(p, d, s)] = self.model.NewBoolVar(
                        f"p{p}_d{d}_s{s}"
                    )

        # Hard constraint 1:
        # Każdy pracownik może pracować maksymalnie na jednej zmianie w jednym dniu
        for p in range(self.num_people):
            for d in range(self.num_days):
                self.model.Add(
                    sum(self.assignments[(p, d, s)] for s in range(self.shifts_per_day)) <= 1
                )

        # Hard constraint 2:
        # Każda zmiana musi mieć dokładnie 1 pracownika dla każdego dnia
        for d in range(self.num_days):
            for s in range(self.shifts_per_day):
                self.model.Add(
                    sum(self.assignments[(p, d, s)] for p in range(self.num_people)) == 1
                )

        # Fairness constraint (soft objective): minimalizacja różnicy między max a min liczbą zmian przypadających na pracownika
        self.person_shift_count = [
            sum(self.assignments[(p, d, s)] for d in range(self.num_days) for s in range(self.shifts_per_day))
            for p in range(self.num_people)
        ]

        self.max_shifts = self.model.NewIntVar(0, self.num_days * self.shifts_per_day, 'max_shifts')
        self.min_shifts = self.model.NewIntVar(0, self.num_days * self.shifts_per_day, 'min_shifts')
        self.model.AddMaxEquality(self.max_shifts, self.person_shift_count)
        self.model.AddMinEquality(self.min_shifts, self.person_shift_count)

        self.model.Minimize(self.max_shifts - self.min_shifts)

        # Optymalizacja preferencji: maksymalizuj przypisania do preferowanych zmian
        self.preference_score = 0
        for p_str, preferred_shifts in self.preferences.items():
            p = int(p_str)  # konwertuj klucz string na int
            if p < self.num_people:  # sprawdź czy pracownik istnieje
                for d in range(self.num_days):
                    for s in preferred_shifts:
                        if s < self.shifts_per_day:
                            self.preference_score += self.assignments[(p, d, s)]
        
        # Dodaj do funkcji celu (maksymalizuj preferencje, minimalizuj nierówność)
        self.model.Maximize(self.preference_score - (self.max_shifts - self.min_shifts))

    def solve(self) -> dict:
        """Rozwiąż model i zwróć dict z wynikiem i przykładowym harmonogramem."""
        solver = cp_model.CpSolver()
        status = solver.Solve(self.model)

        status_text = "OPTIMAL" if status == cp_model.OPTIMAL else (
            "FEASIBLE" if status == cp_model.FEASIBLE else "INFEASIBLE"
        )

        # Przykładowy harmonogram: macierz 14x7 z metadanymi
        schedule = [["" for _ in range(self.num_days)] for _ in range(self.num_people)]

        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            for p in range(self.num_people):
                for d in range(self.num_days):
                    for s in range(self.shifts_per_day):
                        if solver.Value(self.assignments[(p, d, s)]):
                            shift_label = self.shift_names[s] if s < len(self.shift_names) else f"Zmiana {s+1}"
                            # Zamiast prostego stringa, użyj metadanych
                            schedule[p][d] = {
                                "value": shift_label,
                                "color": "#4CAF50" if shift_label == "Poranna" else "#2196F3",
                                "bold": shift_label == "Poranna"
                            }

        response = {
            "status": status_text,
            "schedule": schedule,
            "num_people": self.num_people,
            "num_days": self.num_days,
            "shift_names": self.shift_names,
        }

        # zapis do JSON (Opcja 3)
        try:
            import json
            with open("harmonogram.json", "w", encoding="utf-8") as f:
                json.dump(response, f, ensure_ascii=False, indent=2)
        except Exception as e:
            # nie przerywaj działania, tylko zostaw log (e.g. w przyszłości logger)
            print(f"Błąd zapisu harmonogramu: {e}")

        return response

