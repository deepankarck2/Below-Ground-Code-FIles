"""
This script loads training data from a CSV file, extracts features and labels, trains a RandomForestRegressor model, 
and evaluates the model's performance using Mean Absolute Error (MAE), Mean Squared Error (MSE), Root Mean Squared Error (RMSE), 
and R-squared Value. It also visualizes feature importance. 

The script assumes that the training data is stored in a CSV file with columns representing the features and labels, 
and that the file is located at the specified path. The features are assumed to have names starting with "load" or "gen", 
and the labels are assumed to have names starting with "bus".

The RandomForestRegressor model is initialized with 100 trees and a random state of 42. The data is split into 
80% training and 20% testing sets, and the model is trained on the training set. The performance of the model is 
evaluated on the testing set using MAE, MSE, RMSE, and R-squared Value. The feature importance is also visualized.
"""

from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np


# Load the data
df = pd.read_csv("A:\\CCNY\\J_Fall_2023\\SD2\\OpenDSS\\IEEE 30 Bus\\training_data.csv")

# Extract features and labels
features = df.filter(regex="^(load|gen)").columns
labels = df.filter(regex="^bus").columns

X = df[features]
y = df[labels]

print("Total data points:", len(y))

# Split the data (80% training, 20% testing)
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=42
)

# Initialize the model
model = RandomForestRegressor(n_estimators=100, random_state=42)

# Train the model
model.fit(X_train, y_train)

# Predict
predictions = model.predict(X_test)

# Evaluate the model
mae = mean_absolute_error(y_test, predictions)
mse = mean_squared_error(y_test, predictions)
rmse = np.sqrt(mse)
r2 = r2_score(y_test, predictions)

print(f"Mean Absolute Error (MAE): {mae}")
print(f"Mean Squared Error (MSE): {mse}")
print(f"Root Mean Squared Error (RMSE): {rmse}")
print(f"R-squared Value: {r2}")
print(f"R-squared Value (Accuracy): {r2 * 100:.2f}%")

# # Visualize feature importance
# feature_importances = model.feature_importances_
# sorted_idx = feature_importances.argsort()

# plt.figure(figsize=(10, len(features) * 0.4))
# plt.barh(np.array(features)[sorted_idx], feature_importances[sorted_idx])
# plt.xlabel("Feature Importance")
# plt.title("Feature Importance in RandomForestRegressor")
# plt.show()
