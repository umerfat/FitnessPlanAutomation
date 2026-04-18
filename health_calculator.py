"""
Deterministic health and nutrition calculations.
All formulas are evidence-based and produce consistent, accurate results.
"""


def calculate_bmi(weight_kg: float, height_cm: float) -> dict:
    """Calculate BMI and return value + category."""
    if height_cm <= 0 or weight_kg <= 0:
        raise ValueError(f"Invalid measurements: weight={weight_kg}kg, height={height_cm}cm. Both must be > 0.")
    height_m = height_cm / 100
    bmi = round(weight_kg / (height_m ** 2), 1)

    if bmi < 18.5:
        category = "Underweight"
    elif bmi < 25:
        category = "Normal"
    elif bmi < 30:
        category = "Overweight"
    else:
        category = "Obese"

    return {"bmi": bmi, "category": category}


def calculate_bmr(weight_kg: float, height_cm: float, age: int, gender: str) -> float:
    """Calculate BMR using Mifflin-St Jeor equation."""
    bmr = 10 * weight_kg + 6.25 * height_cm - 5 * age
    if (gender or "male").lower() == "male":
        bmr += 5
    else:
        bmr -= 161
    return round(bmr)


def get_activity_multiplier(activity_level: str) -> float:
    """Map activity level string to TDEE multiplier."""
    level = (activity_level or "").lower()
    if "sedentary" in level:
        return 1.2
    elif "light" in level or "1-3" in level or "1–3" in level:
        return 1.375
    elif "moderate" in level or "3-5" in level or "3–5" in level:
        return 1.55
    elif "very active" in level or "5-6" in level or "5–6" in level or "intense" in level:
        return 1.725
    elif "extra" in level or "athlete" in level:
        return 1.9
    return 1.375  # default to light


def calculate_tdee(bmr: float, activity_level: str) -> float:
    """Calculate Total Daily Energy Expenditure."""
    multiplier = get_activity_multiplier(activity_level)
    return round(bmr * multiplier)


def calculate_target_calories(tdee: float, fitness_goal: str) -> dict:
    """Calculate target calories based on goal."""
    goal = (fitness_goal or "").lower()
    if "fat loss" in goal or "weight loss" in goal or "lose" in goal:
        target = tdee - 500
        strategy = "Caloric deficit (-500 kcal from TDEE)"
    elif "muscle gain" in goal or "bulk" in goal or "mass" in goal:
        target = tdee + 300
        strategy = "Caloric surplus (+300 kcal above TDEE)"
    elif "recomp" in goal or "body recomposition" in goal:
        target = tdee - 150  # slight deficit for recomp
        strategy = "Slight deficit (-150 kcal) for body recomposition"
    else:
        target = tdee
        strategy = "Maintenance calories"

    return {"target_calories": round(target), "strategy": strategy, "tdee": tdee}


def calculate_macros(weight_kg: float, target_calories: float, fitness_goal: str) -> dict:
    """Calculate macro split in grams."""
    goal = (fitness_goal or "").lower()

    # Protein: 1.8-2.2g per kg based on goal
    if "muscle gain" in goal or "bulk" in goal:
        protein_per_kg = 2.2
    elif "fat loss" in goal or "lose" in goal:
        protein_per_kg = 2.0  # higher protein to preserve muscle in deficit
    else:
        protein_per_kg = 2.0

    protein_g = round(weight_kg * protein_per_kg)
    protein_cal = protein_g * 4

    # Fats: 25% of total calories
    fat_cal = target_calories * 0.25
    fat_g = round(fat_cal / 9)

    # Carbs: remaining calories
    carb_cal = target_calories - protein_cal - fat_cal
    carb_g = round(max(carb_cal, 0) / 4)  # floor at 0

    return {
        "protein_g": protein_g,
        "carbs_g": carb_g,
        "fats_g": fat_g,
        "protein_cal": round(protein_cal),
        "carbs_cal": round(carb_cal),
        "fats_cal": round(fat_cal),
    }


def calculate_waist_to_hip_ratio(waist_inches: float, hip_inches: float, gender: str) -> dict:
    """Calculate waist-to-hip ratio (primarily for females)."""
    if not hip_inches or hip_inches == 0:
        return None

    whr = round(waist_inches / hip_inches, 2)
    if (gender or "male").lower() == "female":
        if whr < 0.80:
            risk = "Low health risk"
        elif whr <= 0.85:
            risk = "Moderate health risk"
        else:
            risk = "High health risk"
    else:
        if whr < 0.90:
            risk = "Low health risk"
        elif whr <= 0.95:
            risk = "Moderate health risk"
        else:
            risk = "High health risk"

    return {"whr": whr, "risk": risk}


def calculate_hydration_target(weight_kg: float) -> float:
    """Recommended daily water intake in liters (approx 35ml per kg)."""
    return round(weight_kg * 0.035, 1)


def estimate_body_fat(bmi: float, age: int, gender: str) -> float:
    """Estimate body fat percentage using the Deurenberg formula.
    BF% = 1.20 × BMI + 0.23 × Age − 10.8 × Sex − 5.4
    (Sex: 1 for male, 0 for female)
    """
    sex_factor = 1 if (gender or "male").lower() == "male" else 0
    bf = 1.20 * bmi + 0.23 * age - 10.8 * sex_factor - 5.4
    return round(max(bf, 5.0), 1)  # floor at 5%


def estimate_lean_mass(weight_kg: float, body_fat_pct: float) -> float:
    """Estimate lean body mass from weight and body fat %."""
    return round(weight_kg * (1 - body_fat_pct / 100), 1)


def get_body_fat_category(bf_pct: float, gender: str) -> str:
    """Categorize body fat percentage."""
    if (gender or "male").lower() == "male":
        if bf_pct < 6:
            return "Essential Fat"
        elif bf_pct < 14:
            return "Athletic"
        elif bf_pct < 18:
            return "Fit"
        elif bf_pct < 25:
            return "Average"
        else:
            return "Above Average"
    else:
        if bf_pct < 14:
            return "Essential Fat"
        elif bf_pct < 21:
            return "Athletic"
        elif bf_pct < 25:
            return "Fit"
        elif bf_pct < 32:
            return "Average"
        else:
            return "Above Average"


def compute_all_metrics(client_data: dict) -> dict:
    """Compute all health metrics from client form data."""
    weight = _parse_float(client_data.get("Current Body Weight (in kg)", 0))
    height = _parse_float(client_data.get("Height (in cm)", 0))
    age = int(_parse_float(client_data.get("Age (years)", 0)))
    gender = client_data.get("Gender") or "Male"

    if weight <= 0:
        raise ValueError(f"Invalid weight: '{client_data.get('Current Body Weight (in kg)', '')}'. Must be a positive number.")
    if height <= 0:
        raise ValueError(f"Invalid height: '{client_data.get('Height (in cm)', '')}'. Must be a positive number.")
    if age <= 0 or age > 150:
        raise ValueError(f"Invalid age: '{client_data.get('Age (years)', '')}'. Must be between 1 and 150.")
    waist = _parse_float(client_data.get("Waist Circumference (in inches)", ""))
    hip = _parse_float(client_data.get("Hip Circumference (in inches) (Females only)", ""))
    activity_level = client_data.get("Current Activity Level", "")
    fitness_goal = client_data.get("Primary Fitness Goal", "")

    bmi_data = calculate_bmi(weight, height)
    bmr = calculate_bmr(weight, height, age, gender)
    tdee = calculate_tdee(bmr, activity_level)
    calorie_data = calculate_target_calories(tdee, fitness_goal)
    macro_data = calculate_macros(weight, calorie_data["target_calories"], fitness_goal)
    whr_data = calculate_waist_to_hip_ratio(waist, hip, gender) if hip else None
    hydration = calculate_hydration_target(weight)
    body_fat_pct = estimate_body_fat(bmi_data["bmi"], age, gender)
    lean_mass = estimate_lean_mass(weight, body_fat_pct)
    bf_category = get_body_fat_category(body_fat_pct, gender)

    return {
        "bmi": bmi_data,
        "bmr": bmr,
        "tdee": tdee,
        "calories": calorie_data,
        "macros": macro_data,
        "whr": whr_data,
        "hydration_liters": hydration,
        "body_fat_pct": body_fat_pct,
        "body_fat_category": bf_category,
        "lean_mass_kg": lean_mass,
        "fat_mass_kg": round(weight - lean_mass, 1),
        "weight_kg": weight,
        "height_cm": height,
        "age": age,
        "gender": gender,
    }


def _parse_float(value: str) -> float:
    """Safely parse a float from a string, handling units and empty values."""
    if not value or str(value).strip() == "":
        return 0.0
    # Remove common suffixes like 'kg', 'cm', 'inches'
    cleaned = str(value).lower().replace("kg", "").replace("cm", "").replace("inches", "").strip()
    try:
        return float(cleaned)
    except ValueError:
        return 0.0
