import xml.etree.ElementTree as ET

tree = ET.parse('PRUEBA.xml')

root = tree.getroot()

print(root)
