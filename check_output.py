import json

data = json.load(open(r'c:\Users\Macla\Desktop\AI\dsgl_munitions_list.json', encoding='utf-8'))

ml1 = [x for x in data if x['Label'] == 'ML1'][0]
ml2 = [x for x in data if x['Label'] == 'ML2'][0]
ml3 = [x for x in data if x['Label'] == 'ML3'][0]

print('ML1:', [x['Label'] for x in ml1['SubStructures']])
print('ML2:', [x['Label'] for x in ml2['SubStructures']])
print('ML3:', [x['Label'] for x in ml3['SubStructures']])

ml1_b = [x for x in ml1['SubStructures'] if x['Label'] == 'b.']
print(f"\nML1 has {len(ml1_b)} 'b.' items")

if ml1_b and len(ml1_b[0]['SubStructures']) > 1:
    ml1_b2 = ml1_b[0]['SubStructures'][1]
    print('ML1.b.2 items:', [x['Label'] for x in ml1_b2['SubStructures']])
