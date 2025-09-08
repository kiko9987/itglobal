import requests

response = requests.get('http://localhost:5000')
html = response.text

print("Dashboard Status:", response.status_code)
print("Has refresh button text:", "새로고침" in html)
print("Has refreshData function:", "refreshData" in html)
print("Has refreshBtn ID:", "id=\"refreshBtn\"" in html)

response2 = requests.get('http://localhost:5000/projects')
html2 = response2.text

print("\nProject page Status:", response2.status_code)
print("Has refresh button text:", "새로고침" in html2)
print("Has refreshProjectData function:", "refreshProjectData" in html2)
print("Has refreshProjectBtn ID:", "id=\"refreshProjectBtn\"" in html2)