import mongoose from "mongoose";
// Define a schema for the form data
const formDataSchema = new mongoose.Schema({
  title: { type: String, required: true },
  data: { type: Object, required: true }, // Store the form data as an object
}, { timestamps: true });

// Create a model based on the schema
const FormData = mongoose.model('FormData', formDataSchema);

export default FormData