Tables
======

Tables Test

.. role:: showstopper
.. role:: high
.. role:: medium
.. role:: low


CSV Tables
==========
       
.. csv-table::
   :header: Count, Species

   35, Chimpanzee
   45, Bonobo
   2, Bigfoot

Table with classes
==================
   
.. csv-table::
   :class: issues
   :header: Count, Species


   35, **Chimpanzee**
   45, *Bonobo*
   2, :high:`Bigfoot`

Sample text

Table with widths
=================

.. csv-table::
    :class: rapid
    :widths: 25, 25, 50

    Topic, Description of the proposal, Notes  
    Recommender, Presents a recommendation after assessing relevant input, THis is more notes
