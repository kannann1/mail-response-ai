# test_ollama.py
import requests
import json

def test_ollama_api():
    url = "http://localhost:11434/api/generate"
    
    payload = {
        "model": "llama3",
        "prompt": "Write a short email reply to: Hello, can you help me with a problem?"
    }
    
    print(f"Sending request to {url}")
    print(f"Payload: {json.dumps(payload, indent=2)}")
    
    try:
        response = requests.post(url, json=payload)
        
        print(f"Response status: {response.status_code}")
        print(f"Response headers: {response.headers}")
        
        # Print raw response text
        print(f"Raw response text: {response.text[:500]}...")
        
        # Try to parse the response differently
        try:
            data = response.json()
            print(f"JSON parsed successfully")
        except json.JSONDecodeError as e:
            print(f"JSON parse error: {e}")
            
            # Try to parse line by line
            lines = response.text.strip().split('\n')
            print(f"Response contains {len(lines)} lines")
            
            if len(lines) > 1:
                print("Trying to parse first line...")
                try:
                    first_json = json.loads(lines[0])
                    print(f"First line parsed: {first_json}")
                except:
                    print("Could not parse first line")
            
            return None
        
        return data.get("response", "")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return None

if __name__ == "__main__":
    result = test_ollama_api()
    if result:
        print("\nGenerated text:")
        print(result)
    else:
        print("\nFailed to get a response")