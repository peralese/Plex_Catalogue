from flask import Flask, request, jsonify, render_template
from modules.movie_wishlist_sync import load_movie_wishlist, save_movie_wishlist
import pandas as pd

app = Flask(__name__)
sheet_name = "DVD Wish List"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/wishlist', methods=['GET'])
def get_wishlist():
    df = load_movie_wishlist(sheet_name)
    wishlist = df.fillna('').to_dict(orient='records')
    return jsonify(wishlist)

@app.route('/wishlist', methods=['POST'])
def add_wishlist_item():
    item = request.get_json()
    df = load_movie_wishlist(sheet_name)
    # df = df.append(item, ignore_index=True)
    df = pd.concat([df, pd.DataFrame([item])], ignore_index=True)
    save_movie_wishlist(sheet_name, df)
    return jsonify({"message": "Item added"}), 201

@app.route('/wishlist/<int:index>', methods=['PUT'])
def update_wishlist_item(index):
    item = request.get_json()
    df = load_movie_wishlist(sheet_name)
    if 0 <= index < len(df):
        for key in item:
            if key in df.columns:
                df.at[index, key] = item[key]
        save_movie_wishlist(sheet_name, df)
        return jsonify({"message": "Item updated"}), 200
    else:
        return jsonify({"error": "Invalid index"}), 404

@app.route('/wishlist/<int:index>', methods=['DELETE'])
def delete_wishlist_item(index):
    df = load_movie_wishlist(sheet_name)
    if 0 <= index < len(df):
        df = df.drop(index).reset_index(drop=True)
        save_movie_wishlist(sheet_name, df)
        return jsonify({"message": "Item deleted"}), 200
    else:
        return jsonify({"error": "Invalid index"}), 404

if __name__ == '__main__':
    app.run(debug=True)
