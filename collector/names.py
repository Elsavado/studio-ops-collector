"""Name parts for the fictional company. All people, clients, suppliers and vessels are made up."""

FIRST = """Aisha Omar Fatima Yusuf Layla Khalid Mariam Hassan Noura Tariq Salma Karim Huda Rami Dana Faisal
Priya Arjun Ananya Rahul Kavya Vikram Meera Sanjay Divya Rohan Neha Aditya Lakshmi Imran Sana Farhan
Maria Jose Angelica Mark Kristine Paolo Joy Ramon Grace Miguel Liza Carlo
Wanjiru Kamau Achieng Otieno Njeri Mwangi Akinyi Kiprop Amani Baraka Zawadi Juma Nia Chege
James Emma Daniel Sophie Thomas Olivia Lucas Hannah Peter Claire David Laura Michael Anna Stefan Elena
Chen Wei Li Na Hiroshi Yuki Min Joon Sofia Mateo Ivan Natalia Ahmed Leila Samir Rana Ola Tamer
Michael Jennifer Christopher Jessica Matthew Ashley Joshua Amanda Andrew Brittany Ryan Megan Tyler Lauren
Brandon Kayla Justin Rachel Kevin Nicole Jason Heather Eric Stephanie Kyle Danielle Jacob Courtney Austin
Madison Cody Brianna Dylan Taylor Marcus Jasmine Darnell Tiana Malik Keisha Carlos Gabriela Luis Valeria
Amit Pooja Rajesh Sunita Vivek Shreya Manish Anjali Suresh Deepika Nikhil Ritu Gaurav Swati Harish Kiran
Venkat Bhavna Prakash Neeraj Sneha Abhishek Isha Rakesh Pallavi Arvind Tanvi Siddharth Nandini
Wei Jing Lei Fang Hao Mei Jun Xin Yang Ling Tao Yan Bo Qing Feng Hui Jian Ying Ming Xiu Zhen Lan Kai Rui
Yiming Xiaoling Zhiwei Haoran Yuxuan Zihan Jiahui Shu Peng Yun""".split()

LAST = """Haddad Rahman Al-Amin Farouk Nasser Saleh Khoury Mansour Qasim Darwish Haidar Shaheen Barakat
Sharma Iyer Nair Kapoor Menon Reddy Pillai Desai Bose Joshi Kulkarni Chandra Rao Siddiqui Qureshi
Santos Reyes Cruz Bautista Garcia Mendoza Villanueva Dela Rosa Aquino Ramos Navarro
Odhiambo Wambui Kariuki Mutua Njoroge Ochieng Kimani Wekesa Chebet Mugo Omondi Ndungu
Whitfield Harper Lindqvist Moreau Keller Brennan Novak Fischer Costa Ward Ellison Dunmore Hale
Tanaka Park Zhang Liu Kim Petrov Ivanova Rossi Silva Okafor Mensah Adeyemi Abebe Tesfaye
Smith Johnson Williams Brown Jones Miller Davis Wilson Anderson Thomas Taylor Moore Jackson Martin Lee
Thompson White Harris Clark Lewis Robinson Walker Young Allen King Wright Scott Hill Green Adams Baker
Nelson Carter Mitchell Roberts Turner Phillips Campbell Parker Evans Edwards Collins Stewart Morris
Murphy Cook Rogers Morgan Cooper Peterson Bailey Reed Hernandez Lopez Gonzalez Rodriguez Martinez
Patel Gupta Singh Kumar Verma Agarwal Mehta Shah Malhotra Banerjee Chatterjee Mukherjee Saxena Bhatia
Krishnan Subramanian Venkatesh Shetty Naidu Chopra Trivedi Pandey Mishra Tiwari Dubey Jain Goyal
Wang Li Chen Yang Huang Zhao Wu Zhou Xu Sun Ma Zhu Hu Guo He Lin Gao Luo Zheng Liang Xie Song Tang
Han Feng Deng Cao Peng Zeng Xiao Tian Dong Pan Yuan Cai Jiang Yu Du Ye Cheng Wei Su Lu Ding Shen""".split()

HOTEL_A = """Azure Pearl Coral Saffron Crescent Oasis Falcon Palmcrest Silvertide Marlin Amberbay Sapphire
Desertrose Lagoon Mirage Halcyon Seabreeze Starfall Cedarpoint Moonsand Goldleaf Tidewater Brightwater
Dunecrest Solenne Opalina Zephyr Lumen Aurelia Serenova Kestrel Aldora Velmira Corvina Marisol Ondine
Rivale Sandor Belmare Castora Elyra Fenwick Grayhaven Ivora Juniper Kalina Lirabay Meridia Novella""".split()
HOTEL_B = ["Resort & Spa", "Grand Hotel", "Suites", "Beach Resort", "Palace Hotel", "Marina Hotel",
           "Hotel & Residences", "Island Resort", "Desert Retreat", "City Hotel"]
CRUISE_B = ["Ocean Voyages", "Cruise Line", "Sea Journeys", "Maritime Cruises", "Yacht Cruises"]
PORT_B = ["Marina Operations", "Port Services", "Harbour Authority Services", "Yacht Club"]
TOWER_B = ["Residences", "Towers", "Business Park", "Mall Management"]
SHIP = """Serenity Aurora Horizon Solstice Meridian Tranquil Voyager Celeste Equinox Odyssey Harmony Radiance
Lumina Nautica Azurine Marea Brisa Estrella Coralia Vela Zephyrine Halcyone Oceana Mistral""".split()

# city, country, sales tax / VAT rate on services (approximate, for fictional invoices)
LOCATIONS = [
    ("Dubai", "United Arab Emirates", 0.05), ("Abu Dhabi", "United Arab Emirates", 0.05),
    ("Doha", "Qatar", 0.0), ("Muscat", "Oman", 0.05), ("Riyadh", "Saudi Arabia", 0.15),
    ("Jeddah", "Saudi Arabia", 0.15), ("Manama", "Bahrain", 0.10),
    ("Cancun", "Mexico", 0.16), ("Vancouver", "Canada", 0.05), ("Toronto", "Canada", 0.13),
    ("Nassau", "Bahamas", 0.10), ("San Juan", "Puerto Rico", 0.115),
    ("London", "United Kingdom", 0.20), ("Barcelona", "Spain", 0.21), ("Nice", "France", 0.20),
    ("Venice", "Italy", 0.22), ("Athens", "Greece", 0.24), ("Southampton", "United Kingdom", 0.20),
    ("Mumbai", "India", 0.18), ("Goa", "India", 0.18), ("Bengaluru", "India", 0.18), ("New Delhi", "India", 0.18),
    ("Shanghai", "China", 0.06), ("Sanya", "China", 0.06), ("Hong Kong", "China", 0.0),
    ("Singapore", "Singapore", 0.09), ("Phuket", "Thailand", 0.07), ("Bali", "Indonesia", 0.11),
    ("Tokyo", "Japan", 0.10), ("Sydney", "Australia", 0.10),
    ("Nairobi", "Kenya", 0.16), ("Mombasa", "Kenya", 0.16), ("Zanzibar", "Tanzania", 0.18),
    ("Cape Town", "South Africa", 0.15), ("Mauritius", "Mauritius", 0.15), ("Seychelles", "Seychelles", 0.15),
]
# most clients are in the United States (sales tax on protective services varies by state)
US_LOCATIONS = [
    ("Miami", "United States", 0.07), ("Orlando", "United States", 0.065), ("Fort Lauderdale", "United States", 0.07),
    ("Tampa", "United States", 0.075), ("Las Vegas", "United States", 0.0), ("New York", "United States", 0.08875),
    ("Honolulu", "United States", 0.045), ("Maui", "United States", 0.04166), ("San Diego", "United States", 0.0),
    ("Los Angeles", "United States", 0.0), ("San Francisco", "United States", 0.0), ("Seattle", "United States", 0.0),
    ("Chicago", "United States", 0.0), ("Boston", "United States", 0.0), ("Washington DC", "United States", 0.06),
    ("Atlanta", "United States", 0.0), ("Nashville", "United States", 0.0925), ("New Orleans", "United States", 0.0),
    ("Houston", "United States", 0.0825), ("Dallas", "United States", 0.0825), ("Austin", "United States", 0.0825),
    ("Galveston", "United States", 0.0825), ("Phoenix", "United States", 0.0), ("Scottsdale", "United States", 0.0),
    ("Denver", "United States", 0.0), ("Aspen", "United States", 0.0), ("Charleston", "United States", 0.0),
    ("Savannah", "United States", 0.0), ("Myrtle Beach", "United States", 0.0), ("Atlantic City", "United States", 0.06625),
    ("Philadelphia", "United States", 0.0), ("Baltimore", "United States", 0.0), ("Key West", "United States", 0.075),
    ("Anchorage", "United States", 0.0), ("Juneau", "United States", 0.0), ("Port Canaveral", "United States", 0.07),
]


def location(rng):
    return rng.choice(US_LOCATIONS if rng.random() < 0.7 else LOCATIONS)


NATIONALITIES = ["United States"] * 14 + ["India", "India", "India", "China", "China", "China",
                 "United Kingdom", "Philippines", "Kenya", "UAE", "Egypt", "Mexico", "Brazil", "Germany",
                 "South Africa", "Singapore", "Australia", "Nigeria", "Spain", "Canada", "Japan", "Pakistan"]

SUPPLIERS = [
    ("Northgate Vision Supply", "Equipment & spare parts", "5100"),
    ("Cloudvault Hosting", "Cloud video storage & compute", "5200"),
    ("Globeline Telecom", "Connectivity & SIM data", "5300"),
    ("Meridian Properties", "Office & SOC lease", "5400"),
    ("Swiftfleet Leasing", "Service vehicles", "5500"),
    ("Sentrix Software", "VMS & analytics licences", "5600"),
    ("Anchorpoint Insurance", "Liability & equipment insurance", "5700"),
    ("Brightpath Training", "Training & certification", "6000"),
    ("Ledgerwise Advisory", "Audit & professional fees", "6100"),
    ("Uniforma Workwear", "Uniforms & PPE", "6000"),
]
