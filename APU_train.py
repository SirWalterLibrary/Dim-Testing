import pandas as pd
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error
import joblib

# 1. Load the dataset from your CSV file
# Make sure to adjust the path to where your CSV is saved
df = pd.read_csv('APU_train.csv')

# 2. Define Features (Errors) and Labels (Parameters)
X = df[['error_Length', 'error_Width', 'error_Height']]  # Features: errors
y = df[['a1', 'b1', 'c1', 'd1', 'a2', 'b2', 'c2', 'd2']]  # Labels: parameters

# Split the data into training and testing sets
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Initialize the model
model = RandomForestRegressor(random_state=42)

# Define the grid of hyperparameters to search
param_grid = {
    'n_estimators': [50, 100, 200],
    'max_depth': [None, 10, 20, 30],
    'min_samples_split': [2, 5, 10]
}

# Set up the Grid Search
grid_search = GridSearchCV(estimator=model, param_grid=param_grid, 
                           scoring='neg_mean_squared_error', cv=5)

# Fit the model to the data
grid_search.fit(X_train, y_train)

# Get the best parameters
best_params = grid_search.best_params_
print(f"Best Hyperparameters: {best_params}")

# Evaluate the model with the best parameters
best_model = grid_search.best_estimator_
y_pred = best_model.predict(X_test)
mse = mean_squared_error(y_test, y_pred)
print(f"Mean Squared Error with best hyperparameters: {mse:.4f}")

# 8. Save the trained model to a file for future use
joblib.dump(best_model, 'parameter_prediction_model.pkl')

# 9. Load the saved model (optional, to reuse later)
loaded_model = joblib.load('parameter_prediction_model.pkl')

# 10. Test the model by predicting parameter values for new error values
new_errors = df[['error_Length', 'error_Width', 'error_Height']]  # Replace with your actual new errors
predicted_parameters = loaded_model.predict(new_errors)

# Define the parameter names
parameter_names = ['a1', 'b1', 'c1', 'd1', 'a2', 'b2', 'c2', 'd2']

# Format and print predicted parameters in 0.0000 format
formatted_output = []
for name, param in zip(parameter_names, predicted_parameters[0]):  # Iterate over names and values
    formatted_output.append(f"{name} = {param:.4f}")  # Create the formatted string

# Join and print the output with an additional newline for space
print(f"Predicted Parameters for new errors:\n" + "\n".join(formatted_output[:4]) + "\n")  # First four parameters
print("\n".join(formatted_output[4:]))  # Next four parameters
