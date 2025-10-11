# iris.py
# Classify iris flowers using K-Nearest Neighbors

from sklearn.datasets import load_iris
from sklearn.model_selection import train_test_split
from sklearn.neighbors import KNeighborsClassifier
from sklearn.metrics import accuracy_score

# Load dataset
iris = load_iris()
X, y = iris.data, iris.target

# Split dataset
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Train model
model = KNeighborsClassifier(n_neighbors=3)
model.fit(X_train, y_train)

# Predict
y_pred = model.predict(X_test)

# Evaluate
print("Accuracy:", accuracy_score(y_test, y_pred))

# Example prediction with actual flower name
example = [[5.1, 3.5, 1.4, 0.2]]
pred_class = model.predict(example)[0]
pred_name = iris.target_names[pred_class]
print("Example input:", example)
print("Predicted Flower:", pred_name)
