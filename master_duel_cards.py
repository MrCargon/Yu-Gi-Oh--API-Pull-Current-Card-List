import requests
import pandas as pd
import time

def fetch_master_duel_cards():
    print("Starting to fetch Master Duel cards...")
    base_url = "https://db.ygoprodeck.com/api/v7/cardinfo.php"
    params = {
        "format": "Master Duel",
        "misc": "yes"
    }
    
    all_cards = []
    
    try:
        print("Sending API request...")
        response = requests.get(base_url, params=params)
        response.raise_for_status()
        data = response.json()
        print(f"Received data for {len(data['data'])} cards.")
        
        for index, card in enumerate(data['data']):
            if index % 1000 == 0:
                print(f"Processing card {index + 1}...")
            card_info = {
                'id': card['id'],
                'name': card['name'],
                'type': card['type'],
                'desc': card['desc'],
                'race': card['race'],
                'archetype': card.get('archetype', ''),
                'card_sets': ', '.join([set['set_name'] for set in card.get('card_sets', [])]),
                'ban_status': card.get('banlist_info', {}).get('ban_master_duel', 'Unlimited')
            }
            
            if 'atk' in card:
                card_info['atk'] = card['atk']
            if 'def' in card:
                card_info['def'] = card['def']
            if 'level' in card:
                card_info['level'] = card['level']
            elif 'linkval' in card:
                card_info['link'] = card['linkval']
            
            all_cards.append(card_info)
        
        print(f"Processed {len(all_cards)} cards.")
        return all_cards
    except requests.RequestException as e:
        print(f"An error occurred during API request: {e}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None

def save_to_excel(cards, filename='master_duel_cards.xlsx'):
    print(f"Saving {len(cards)} cards to Excel...")
    df = pd.DataFrame(cards)
    df.to_excel(filename, index=False)
    print(f"Card data saved to {filename}")

if __name__ == "__main__":
    print("Script started.")
    print("Fetching Master Duel card data...")
    cards = fetch_master_duel_cards()
    
    if cards:
        print(f"Retrieved {len(cards)} cards.")
        save_to_excel(cards)
    else:
        print("Failed to retrieve card data.")
    
    print("Script finished.")