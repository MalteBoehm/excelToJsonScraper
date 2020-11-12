import xlrd
from collections import OrderedDict
import simplejson as json


### Settings ###
# where to splitt text in case of synonyms
whatSymbolToSplit = "|"

# open excel
wb = xlrd.open_workbook('excelInput/inputfile.xlsx')
# sheet tab <number>
sheet = wb.sheet_by_index(2)


allergensList = []

def addFoodCategoryToAllergenList(deFoodCategoryName, engFoodCategoryName,
                                  deGroupOfFood, engGroupOfFood,
                                  fromRow, toRow) -> []:

    for rownum in range(fromRow, toRow):
        allergens = OrderedDict()
        row_values = sheet.row_values(rownum)
        #Double Assigned to save space
        allergens['names'] = convertsNames(row_values[9]),
        allergens['englishNames']  = convertsNames(row_values[10])
        allergens['deFoodCategory'] = deFoodCategoryName
        allergens['engFoodCategory'] = engFoodCategoryName
        allergens['deSpecificFood'] = deGroupOfFood
        allergens['engSpecificFood'] = engGroupOfFood
        allergens['laktose']  = row_values[2]
        allergens['gluten'] = row_values[3]
        allergens['histamin'] = row_values[4]
        allergens['histaminWirkung'] = row_values[5]
        allergens['weitereArmineWirkung'] = row_values[6]
        allergens['liberatorWirkung'] = row_values[7]
        allergens['blockerWirkung'] = row_values[8]
        allergensList.append(allergens)


def convertsNames(names):
    return names.title().split(whatSymbolToSplit)


# Categorys to Add


### Tierisch / Animal foods ###
# Eier / Eggs
addFoodCategoryToAllergenList("Tierisch", "Animal foods",
                              "Eier", "Eggs",
                              3, 7)
# Milchprodukte
addFoodCategoryToAllergenList("Tierisch", "Animal foods",
                              "Milchprodukte", "Dairy products",
                              9, 53)
# Fleisch
addFoodCategoryToAllergenList("Tierisch", "Animal foods",
                              "Fleisch", "Meat",
                              55, 84)
# Fisch
addFoodCategoryToAllergenList("Tierisch", "Animal foods",
                              "Fisch", "Fish",
                              85, 90)
# Meeresfrüchte
addFoodCategoryToAllergenList("Tierisch", "Animal foods",
                              "Meeresfrüchte", "Sea food",
                              92, 102)
# Diverses
addFoodCategoryToAllergenList("Tierisch", "Animal foods",
                              "Diverses", "Miscellaneous",
                              104, 104)

###  Pflanzlich / Vegetable foods ###
# Stärkelieferanten
addFoodCategoryToAllergenList("Pflanzlich", "Vegetable foods",
                              "Stärkelieferanten", "Starch suppliers",
                              107, 150)
# Nüsse
addFoodCategoryToAllergenList("Pflanzlich", "Vegetable foods",
                              "Nüsse", "Nuts",
                              152, 166)
# Fette und Öle
addFoodCategoryToAllergenList("Pflanzlich", "Vegetable foods",
                              "Fette und Öle", "Fats and oils",
                              168, 182)
# Gemüse und Öle
addFoodCategoryToAllergenList("Pflanzlich", "Vegetable foods",
                              "Gemüse", "Vegetables",
                              184, 294)
# Küchenkräuter
addFoodCategoryToAllergenList("Pflanzlich", "Vegetable foods",
                              "Küchenkräuter", "Herbs",
                              296, 308)

# Früchte
addFoodCategoryToAllergenList("Pflanzlich", "Vegetable foods",
                              "Früchte", "Fruits",
                              310, 406)
# Samen
addFoodCategoryToAllergenList("Pflanzlich", "Vegetable foods",
                              "Samen", "Seeds",
                              408, 411)
### Pilze und Algen ###
addFoodCategoryToAllergenList("Pilze und Algen", "Mushrooms, fungi and algae",
                              "Pilze und Algen", "Mushrooms, fungi and algae",
                              413, 431)

### Süssungsmittel ###
addFoodCategoryToAllergenList("Süssungsmittel", "Sweeteners",
                              "Süssungsmittel", "Sweeteners",
                              433, 464)
### Würzen, Gewürze ###
addFoodCategoryToAllergenList("Würzen, Gewürze", "Spices, seasoning, aroma",
                              "Würzen, Gewürze", "Spices, seasoning, aroma",
                              466, 510)

### Getränke / Beverages ###
# Wasser / Water
addFoodCategoryToAllergenList("Getränke", "Beverages",
                              "Wasser", "Water",
                              513, 515)
# Alkoholhaltiges
addFoodCategoryToAllergenList("Getränke", "Beverages",
                              "Alkoholhaltiges", "Alcoholic",
                              517, 533)
# Tee
addFoodCategoryToAllergenList("Getränke", "Beverages",
                              "Tee", "Herbal infusions",
                              535, 546)
# Fruchtsäfte, Nektare
addFoodCategoryToAllergenList("Getränke", "Beverages",
                              "Fruchtsäfte, Nektare", "Juices, fruit nectars",
                              548, 549)
# Gemüsesäfte
addFoodCategoryToAllergenList("Getränke", "Beverages",
                              "Gemüsesäfte", "Vegetable juices",
                              551, 551)
# Koffeingetränke
addFoodCategoryToAllergenList("Getränke", "Beverages",
                              "Koffeingetränke", "Drinks containing coffeine",
                              553, 557)
# Milchersatz
addFoodCategoryToAllergenList("Getränke", "Beverages",
                              "Milchersatz", "Milk surrogates",
                              559, 561)
# Süssgetränke, Limonaden
addFoodCategoryToAllergenList("Getränke", "Beverages",
                              "Süssgetränke, Limonaden", "Soft drinks, soda",
                              563, 568)

### Zusatzstoffe ###
addFoodCategoryToAllergenList("Zusatzstoffe", "Food additives",
                              "Zusatzstoffe", "Food additives",
                              570, 957)
### Vitamine, Mineralstoffe, Spurenelemente, Stimulantien ###
addFoodCategoryToAllergenList("Vitamine, Mineralstoffe, Spurenelemente, Stimulantien",
                              "Vitamins, dietary minerals, trace elements, stimulants",
                              "Vitamine, Mineralstoffe, Spurenelemente, Stimulantien",
                              "Vitamins, dietary minerals, trace elements, stimulants",
                              959, 966)

### Zubereitungen ###
addFoodCategoryToAllergenList("Zubereitungen", "Preparations, mixtures",
                              "Zubereitungen", "Preparations, mixtures",
                              968, 974)

### IN JSON mit UniCode SPEICHERN ###
j = json.dumps(allergensList, ensure_ascii=False)

with open('data.json', 'w', encoding="utf-8") as f:
    f.write(j)
