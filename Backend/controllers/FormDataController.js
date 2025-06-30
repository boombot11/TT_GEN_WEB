
import FormData from '../config/FormModel.js';

const getFormData = async (req, res) => {
    try {
      // Find all documents in the "formData" collection, only select title and data fields
      const formDatas = await FormData.find({}, 'title data');
  
      // Return the result as an array
      return res.status(200).json(formDatas);
    } catch (error) {
      console.error('Error retrieving form data:', error);
      return res.status(500).json({ message: 'Error retrieving form data', error: error.message });
    }
  };

  const getFormSpecificData = async (req, res) => {
    try {
        const  {title} = req.body;
        console.log('entering function');
        console.log(title)
      // Find all documents in the "formData" collection, only select title and data fields
      const formDatas = await FormData.findOne({title});
  
      // Return the result as an array
      return res.status(200).json(formDatas);
    } catch (error) {
      console.error('Error retrieving form data:', error);
      return res.status(500).json({ message: 'Error retrieving form data', error: error.message });
    }
  };

const saveFormData= async (req, res) => {
  const { title, formData } = req.body;

  if (!title || !formData) {
    return res.status(400).json({ message: 'Title and formData are required' });
  }

  try {
    // Create a new document and save it to the database
    const newFormData = new FormData({
      title,
      data: formData, // formData is the actual object containing form data
    });

    await newFormData.save();
    return res.status(201).json({ message: 'Form data stored successfully', newFormData });
  } catch (error) {
    console.error('Error storing form data:', error);
    return res.status(500).json({ message: 'Error storing form data', error: error.message });
  }
};

const deleteFormData = async (req, res) => {
  try {
    const { title } = req.body; // Extract title from the request body

    if (!title) {
      return res.status(400).json({ message: 'Title is required' });
    }

    // Find and delete the document with the given title
    const result = await FormData.findOneAndDelete({ title });

    if (!result) {
      return res.status(404).json({ message: 'Config not found' });
    }

    return res.status(200).json({ message: 'Config deleted successfully', result });
  } catch (error) {
    console.error('Error deleting form data:', error);
    return res.status(500).json({ message: 'Error deleting form data', error: error.message });
  }
};

export const deleteForm = deleteFormData;
export const saveForm=saveFormData;
export const getForm=getFormData;
export const getFormSpecific=getFormSpecificData;