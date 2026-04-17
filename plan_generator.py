"""
Generate personalized fitness plans using Google Gemini API.
Receives pre-calculated health metrics + client data, returns structured plan.
"""

import json

from google import genai

import config


def generate_plan(client_data: dict, metrics: dict) -> dict:
    """
    Send client data and pre-calculated metrics to Gemini.
    Returns a structured plan dict with nutrition, training, and lifestyle sections.
    """
    prompt = _build_prompt(client_data, metrics)

    client = genai.Client(api_key=config.GEMINI_API_KEY)
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
    )

    response_text = response.text

    # Extract JSON from response (Gemini may wrap it in markdown code blocks)
    json_str = _extract_json(response_text)
    plan = json.loads(json_str)
    return plan


def _build_prompt(client_data: dict, metrics: dict) -> str:
    """Build the detailed prompt for Claude."""

    # Client profile summary
    name = client_data.get("Full Name", "Client")
    age = metrics["age"]
    gender = metrics["gender"]
    weight = metrics["weight_kg"]
    height = metrics["height_cm"]
    goal = client_data.get("Primary Fitness Goal", "General fitness")
    motivation = client_data.get("Why do you want this transformation?", "")
    activity_level = client_data.get("Current Activity Level", "")
    experience = client_data.get("Training Experience", "Beginner")
    gym_access = client_data.get("Gym Access", "Home workouts")
    sleep = client_data.get("Average Sleep per Night", "")
    phone_usage = client_data.get("Late-night phone usage after 10 pm", "")
    stress = client_data.get("Daily Stress Level", "")
    occupation = client_data.get("Occupation Type", "")
    diet_pref = client_data.get("Diet Preference", "Non-vegetarian")
    whey_ok = client_data.get("Is whey protein acceptable?", "Yes")
    allergies = client_data.get("Any food allergies or foods you dislike?", "No")
    injuries = client_data.get("Any injuries or medical conditions?", "No")

    # Pre-calculated metrics
    bmi = metrics["bmi"]
    bmr = metrics["bmr"]
    tdee = metrics["tdee"]
    cal = metrics["calories"]
    macros = metrics["macros"]
    whr = metrics.get("whr")
    hydration = metrics["hydration_liters"]
    body_fat = metrics["body_fat_pct"]
    bf_category = metrics["body_fat_category"]
    lean_mass = metrics["lean_mass_kg"]
    fat_mass = metrics["fat_mass_kg"]

    whr_text = ""
    if whr:
        whr_text = f"- Waist-to-Hip Ratio: {whr['whr']} ({whr['risk']})"

    prompt = f"""You are an expert fitness and nutrition coach creating a personalized transformation protocol for a client.

## CLIENT PROFILE
- Name: {name}
- Age: {age}
- Gender: {gender}
- Weight: {weight} kg
- Height: {height} cm
- Occupation: {occupation}
- Primary Goal: {goal}
- Motivation: {motivation}

## LIFESTYLE DATA
- Current Activity Level: {activity_level}
- Training Experience: {experience}
- Gym Access: {gym_access}
- Average Sleep: {sleep}
- Late-night Phone Usage: {phone_usage}
- Daily Stress Level: {stress}

## DIETARY PREFERENCES
- Diet Type: {diet_pref}
- Whey Protein Acceptable: {whey_ok}
- Allergies / Disliked Foods: {allergies}

## MEDICAL
- Injuries / Conditions: {injuries}

## PRE-CALCULATED HEALTH METRICS (use these exact values, do NOT recalculate)
- BMI: {bmi['bmi']} ({bmi['category']})
- Estimated Body Fat: {body_fat}% ({bf_category})
- Lean Body Mass: {lean_mass} kg
- Fat Mass: {fat_mass} kg
- BMR: {bmr} kcal
- TDEE: {tdee} kcal
- Target Calories: {cal['target_calories']} kcal ({cal['strategy']})
- Macros: Protein {macros['protein_g']}g | Carbs {macros['carbs_g']}g | Fats {macros['fats_g']}g
{whr_text}
- Recommended Hydration: {hydration}L/day

## INSTRUCTIONS

Generate a complete transformation protocol. Return ONLY valid JSON (no markdown, no code fences) with this exact structure:

{{
  "physique_insight": "A 3-4 sentence professional assessment. Reference the exact BMI ({bmi['bmi']}), body fat ({body_fat}%), lean mass ({lean_mass}kg), and fat mass ({fat_mass}kg). Explain what these numbers mean for the client's goal. Be encouraging but honest.",

  "nutrition_plan": {{
    "meals": [
      {{
        "timing": "Morning Breakfast",
        "time_suggestion": "7:00 AM",
        "items": "Primary meal: specific foods with quantities",
        "alternative": "Alternative option: equally nutritious swap with similar macros",
        "protein": <grams as number>,
        "carbs": <grams as number>,
        "fats": <grams as number>
      }}
    ],
    "rules": ["Rule 1", "Rule 2", "..."]
  }},

  "training_split": {{
    "days": [
      {{
        "day": "Monday",
        "focus": "Chest",
        "exercises": [
          {{
            "name": "Exercise Name",
            "sets_reps": "4x8",
            "rest": "90s",
            "alternative": "Alternative exercise name with same sets/reps",
            "notes": "brief form tip"
          }}
        ]
      }}
    ]
  }},

  "lifestyle_upgrade": ["Tip 1", "Tip 2", "..."],

  "supplements": ["Supplement 1 with dosage", "..."]
}}

## PLAN GENERATION RULES

### Nutrition:
1. Create exactly 7 meals/snacks spread across the day.
2. Meal timings should be realistic for the client's occupation ({occupation}).
3. Per-meal macros MUST sum up closely to the targets: Protein {macros['protein_g']}g, Carbs {macros['carbs_g']}g, Fats {macros['fats_g']}g (within ±5g tolerance).
4. {"ONLY use vegetarian foods (no meat, no fish, no eggs — paneer, dal, tofu, soy are fine). Eggs are NOT vegetarian." if diet_pref.lower() == "vegetarian" else "Include a mix of non-vegetarian and vegetarian meals."}
5. {"Do NOT include whey protein in any meal." if whey_ok.lower() != "yes" else "Whey protein can be included post-workout."}
6. {"EXCLUDE these foods completely: " + allergies if allergies.lower() not in ("no", "none", "") else "No known allergies."}
7. Use common, affordable Indian/South Asian foods (roti, rice, dal, chicken, paneer, oats, eggs, etc.) unless the client's profile suggests otherwise.
8. Include at least 3-4 nutrition rules personalized to the client's goal and lifestyle.
9. IMPORTANT: For EVERY meal, provide an "alternative" field with a realistic swap option. The alternative should have similar macros but use different, equally affordable ingredients. For example, if primary has walnuts, alternative could use peanuts or flaxseeds.

### Training:
1. Create a 5-day training split (Monday to Friday), weekends off.
2. {"Adapt exercises for HOME WORKOUTS only — bodyweight, dumbbells, resistance bands. No barbell or cable machine exercises." if "home" in gym_access.lower() else "Use full gym equipment — barbells, dumbbells, cables, machines."}
3. {"CRITICAL: Client has the following condition — " + injuries + ". Avoid exercises that aggravate this. Suggest safe alternatives." if injuries.lower() not in ("no", "none", "") else "No injury restrictions."}
4. For experience level '{experience}': {"Focus on progressive overload, compound lifts, moderate-heavy weights." if "intermediate" in experience.lower() else "Focus on form, lighter weights, full-body or simple splits." if "beginner" in experience.lower() else "Include advanced techniques like drop sets, supersets, periodization."}
5. Include 4-5 exercises per day.
6. Specify rest periods between sets (shorter for fat loss/endurance, longer for strength/muscle gain).
7. IMPORTANT: For EVERY exercise, provide an "alternative" field with an equally effective substitute exercise that targets the same muscle group. This gives clients flexibility.

### Lifestyle:
1. Provide 4-6 personalized lifestyle tips based on the client's sleep ({sleep}), stress ({stress}), phone usage ({phone_usage}), and occupation ({occupation}).
2. Include specific, actionable advice — not generic platitudes.

### Supplements:
1. Recommend 2-4 supplements appropriate for the goal. Include dosage.
2. {"Do NOT recommend whey protein." if whey_ok.lower() != "yes" else ""}

Return ONLY the JSON object. No explanations before or after."""

    return prompt


def _extract_json(text: str) -> str:
    """Extract JSON from Claude's response, handling possible markdown wrapping."""
    text = text.strip()

    # Remove markdown code fences if present
    if text.startswith("```"):
        lines = text.split("\n")
        # Remove first line (```json or ```) and last line (```)
        start = 1
        end = len(lines) - 1
        if lines[end].strip() == "```":
            text = "\n".join(lines[start:end])
        else:
            text = "\n".join(lines[start:])

    # Find the JSON object boundaries
    first_brace = text.find("{")
    last_brace = text.rfind("}")
    if first_brace != -1 and last_brace != -1:
        text = text[first_brace : last_brace + 1]

    return text.strip()
