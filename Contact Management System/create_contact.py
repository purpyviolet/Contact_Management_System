
import random
from address_book import add_contact


def create_mixed_contacts(num_objects):
    """
    Creates a specified number of mixed type contacts (Student, Colleague, Friend, Relative)
    with partially similar information.

    :param num_objects: Number of objects to create
    :return: List of created objects of mixed types
    """
    contacts = []
    names = [
    "Emma", "Liam", "Olivia", "Noah", "Ava", "Isabella", "Sophia", "Mia", "Charlotte", "Amelia",    "Harper", "Evelyn", "Abigail", "Emily", "Elizabeth", "Sofia", "Avery", "Ella", "Scarlett", "Grace",
    "Chloe", "Victoria", "Madison", "Aria", "Lily", "Zoe", "Penelope", "Riley", "Layla", "Lucy",    "Aubrey", "Brooklyn", "Harper", "Samantha", "Natalie", "Hannah", "Grace", "Amelia", "Lillian", "Leah",
    "Zoey", "Hazel", "Violet", "Aurora", "Savannah", "Addison", "Bella", "Skylar", "Lily", "Eleanor",    "Stella", "Claire", "Grace", "Savannah", "Emily", "Madison", "Scarlett", "Sophie", "Olivia", "Camila",
    "Eleanor", "Leah", "Hannah", "Sofia", "Mia", "Elizabeth", "Avery", "Ella", "Layla", "Penelope",    "Aubrey", "Chloe", "Lily", "Zoey", "Victoria", "Addison", "Samantha", "Natalie", "Harper", "Brooklyn",
    "Scarlett", "Grace", "Chloe", "Penelope", "Victoria", "Zoey", "Aubrey", "Samantha", "Avery", "Lillian",    "Hannah", "Emily", "Mia", "Layla", "Scarlett", "Addison", "Stella", "Bella", "Claire", "Lily"
]
    types = ["student", "colleague", "friend", "relative"]
    
    relationships = [
    "father", "mother", "son", "daughter", "brother", "sister", "older brother", "older sister","grandfather", "grandmother", "grandson", "granddaughter", "uncle", "aunt", "uncle (maternal)", "aunt (maternal)","nephew", "niece", "cousin (male)", "cousin (female)"
]
    colleges =  [
    "Harvard University", "Stanford University", "Massachusetts Institute of Technology", "California Institute of Technology", "University of Oxford",
    "University of Cambridge", "Princeton University", "Yale University", "Columbia University", "University of Chicago",
    "University of California, Berkeley", "Imperial College London", "University of Toronto", "University of Michigan, Ann Arbor", "University of California, Los Angeles",
    "University of Pennsylvania", "Cornell University", "Swiss Federal Institute of Technology Zurich", "National University of Singapore", "University of Texas at Austin",
    "University of Washington", "University of California, San Diego", "University of Edinburgh", "University of Melbourne", "University of Sydney",
    "University of Hong Kong", "University of Tokyo", "University of Copenhagen", "University of Amsterdam", "University of Munich",
    "Peking University", "Tsinghua University", "Seoul National University", "University of São Paulo", "University of Cape Town",
    "University of Buenos Aires", "University of Toronto", "University of Zurich", "University of Helsinki", "University of Vienna",
    "University of Oslo", "University of Warsaw", "University of Madrid", "University of Milan", "University of Stockholm",
    "University of Delhi", "University of Mumbai", "University of Sao Paulo", "University of Sydney", "University of Queensland"
]

    companies =  [
    "Techtronics Inc.", "InnovateTech Solutions", "GlobalTech Systems", "DataWave Technologies", "Futurix Corporation",
    "EcoTech Innovations", "NexaSoft Solutions", "StrategicSys Group", "Quantum Innovations", "Infotech Solutions",
    "PowerTech Industries", "MegaSoft Technologies", "PinnacleTech Group", "SynergiCorp", "DigitalWave Enterprises",
    "Synthronix Solutions", "InfoSynergy Systems", "CyberCraft Technologies", "InfiniteLoop Innovations", "eFusion Solutions",
    "FusionPoint Systems", "InnoVista Technologies", "MatrixStream Corp.", "NovelSys Solutions", "TechPulse Innovations",
    "DataSphere Systems", "InnoGenix Corporation", "TechNexa Solutions", "StratosWave Tech", "InfoLink Innovations",
    "TechFusion Group", "NexaCore Technologies", "GlobalView Innovations", "ByteCraft Systems", "FusionDynamics Inc.",
    "MetaWave Solutions", "QuantumSys Group", "DigitalEdge Innovations", "TechMasters Corp.", "InfoNexa Technologies",
    "Futurisys Group", "NexaWave Innovations", "Synthronix Corp.", "InnoStream Solutions", "TechVista Systems",
    "InnoMatrix Innovations", "InfoPulse Technologies", "eFusion Systems", "FusionMatrix Innovations", "TechCore Corp."
]


    for _ in range(num_objects):
        contact_type = random.choice(types)
        name = random.choice(names)
        birthday = f"198{random.randint(0, 9)}-{random.randint(1, 12):02d}-{random.randint(1, 28):02d}"
        phone = f"{random.randint(100, 999)}-{random.randint(100, 999)}-{random.randint(1000, 9999)}"
        email = f"{name.lower()}@example.com"
        height = random.randint(150, 200)
        weight = random.randint(50, 100)

        if contact_type == "student":
            college = random.choice(colleges)
            grade = random.choice(["Freshman", "Sophomore", "Junior", "Senior"])
            major = random.choice(["Computer Science", "Physics", "Engineering"])
            gpa = round(random.uniform(2.0, 4.0), 2)
            add_contact( "student",name, birthday, phone, email, height, weight, [college, grade, major, gpa])
        
        elif contact_type == "colleague":
            company = random.choice(companies)
            income = random.randint(50000, 120000)
            add_contact("colleague", name, birthday, phone, email, height, weight, [company, income])

        elif contact_type == "friend":
            acquaintance_info = random.choice(["Met at a conference", "Childhood friend", "University roommate"])
            add_contact("friend",name, birthday, phone, email, height, weight, [acquaintance_info])

        elif contact_type == "relative":
            relationship = random.choice(relationships)
            add_contact( "relative",name, birthday, phone, email, height, weight, [relationship])
