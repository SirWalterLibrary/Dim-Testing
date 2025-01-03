import os
import joblib
import datetime
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.neighbors import KNeighborsClassifier

# Load the training dataset
training = pd.read_csv('data/Xtrain.csv')

# Assume the CSV file has columns: 'length', 'width', 'height', 'label'
X = training[['Length', 'Width', 'Height']]
y = training['Box']

# Split the data into training and testing sets
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Create and train a KNN model
model = KNeighborsClassifier(n_neighbors=3)
model.fit(X_train, y_train)

# Evaluate the model's accuracy
accuracy = model.score(X_test, y_test)
print(f'Model Accuracy: {accuracy*100:.2f}%')

# Get the current date and time
now = datetime.datetime.now()
timestamp = now.strftime("%Y-%m-%d-%H%M")

# Create a directory to store HTML files
output_dir = 'models'
os.makedirs(output_dir, exist_ok=True)

# Output model with timestamp
output_file = os.path.join(output_dir, f'model_{timestamp}.joblib')
joblib.dump(model, output_file)
os.startfile(output_dir)