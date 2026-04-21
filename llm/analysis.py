#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LLMMixin — LLM-powered analysis methods for SyllabusChecker.
"""

import os
import sys
import json
from typing import Dict


def _get_openai():
    """Return OpenAI class, resolving through syllabus_checker so mocks work."""
    # Tests patch 'syllabus_checker.OpenAI'; resolve through that module when
    # available so that mocker.patch('syllabus_checker.OpenAI') still intercepts
    # calls made from within these mixin methods.
    sc = sys.modules.get('syllabus_checker')
    if sc is not None and hasattr(sc, 'OpenAI'):
        return sc.OpenAI
    from openai import OpenAI
    return OpenAI


class LLMMixin:
    """Mixin providing LLM-powered analysis methods."""

    def analyze_growth_mindset_and_belonging(self) -> Dict[str, str]:
        """
        Use OpenAI GPT API to analyze syllabus for growth mindset and belonging cues.
        Returns a dictionary with analysis for each of the 6 core questions.
        """
        # Extract full syllabus text
        syllabus_text = []

        # Get text from all paragraphs
        for para_info in self.all_paragraphs:
            text = para_info.paragraph.text.strip()
            if text:
                syllabus_text.append(text)

        # Join all text
        full_text = "\n\n".join(syllabus_text)

        # Get API key from environment
        api_key = os.environ.get('OPENAI_API_KEY')
        if not api_key:
            return {
                'error': 'OPENAI_API_KEY environment variable not set. Please configure your API key in .env file to enable growth mindset analysis.',
                'status': 'error'
            }

        try:
            client = _get_openai()(api_key=api_key)

            # Create comprehensive prompt based on the provided guidance
            prompt = f"""You are an expert in educational equity and student-centered course design. Your task is to review a course syllabus and evaluate how well it communicates growth mindset and promotes student belonging.

    SYLLABUS TEXT:
    {full_text}

    Please analyze this syllabus according to the following 6 core questions. For each question, provide:
    1. A brief assessment (Strong / Moderate / Weak / Not Addressed)
    2. Specific evidence from the syllabus (quote relevant passages or note absence)
    3. Concrete suggestions for improvement with example language
    4. Ready-to-use text that can be copy-pasted directly into the syllabus, with clear guidance on WHERE to place it

    **CORE QUESTION 1: Growth Mindset**
    Does this syllabus communicate that the instructor has a "growth mindset" rather than a "fixed mindset" about students' abilities?

    Growth mindset conveys that students can grow their abilities through effort, strategies, and help-seeking. Look for:
    - Language about developing skills and abilities over time
    - Assurances that students can succeed with effort and support
    - References to learning as a process vs. innate talent
    - Avoid fixed mindset language like "this course is only for quick learners"

    **CORE QUESTION 2: Normalizing Challenge**
    Do the messages communicate that it is normal to be challenged by course material, and that this is not a sign that a student is not capable or does not belong?

    Look for:
    - Acknowledgment that students may struggle or face setbacks
    - Framing challenges as part of learning, not indicators of inability
    - Instructor/TA stories about overcoming their own challenges
    - Guidance on what to do when struggling (e.g., seek help, use resources)

    **CORE QUESTION 3: Instructor Care**
    Does the syllabus communicate that the instructor and instructional team care about students' success?

    Look for:
    - Personal language about caring for student success
    - Accessible office hours (called "student drop-in hours")
    - Responsive email policies
    - Personal info about TAs to make them more approachable
    - Acknowledgment of large class sizes with assurances of care

    **CORE QUESTION 4: Valuing Diversity**
    Does the syllabus communicate that diversity is valued in the classroom?

    Look for:
    - Land acknowledgments
    - Inclusive course conduct policies
    - Recognition of diverse student backgrounds and experiences
    - Caregiver/family responsibility policies
    - Respectful disagreement and civil discourse expectations
    - Statements about creating welcoming environment for all

    **CORE QUESTION 5: Normalizing Student Challenges**
    Does the syllabus normalize challenges that students often face in college and connect students with resources?

    Look for:
    - Acknowledgment of financial strain (textbook affordability, food insecurity)
    - Mental health and counseling services presented without stigma
    - Disability/accessibility services normalized
    - Work-life balance and nontraditional student support
    - Resources presented as standard tools for success, not remedial

    **CORE QUESTION 6: Normalizing Academic Support Use**
    Does the syllabus communicate that utilizing academic resources is a standard part of succeeding?

    Look for:
    - Office hours framed as "student drop-in hours"
    - Tutoring/academic support presented as normal, not remedial
    - Multiple access points (in-person, virtual, flexible timing)
    - Course chat/discussion forums for questions
    - Language like "most students access these supports"

    IMPORTANT FORMATTING INSTRUCTIONS:
    - Structure your response with clear headings for each of the 6 core questions
    - Use this exact format for each section:

    ## QUESTION [NUMBER]: [Brief Title]
    **Assessment:** [Strong/Moderate/Weak/Not Addressed]

    **Evidence:**
    [Quote specific passages or note absence]

    **Suggestions:**
    [Provide 2-3 concrete, actionable improvements with example language]

    **Ready-to-Use Text:**
    **Where to place it:** [Specify the section of the syllabus - e.g., "Add a new section titled 'Student Support Resources' after the grading policy" or "Insert at the end of the 'Course Description' section" or "Create a new 'Campus Resources' section near the end of the syllabus"]

    **Text to insert:**
    ```
    [Provide 2-4 paragraphs of polished, ready-to-use text that the instructor can copy and paste directly into their syllabus. This should be written in first person from the instructor's perspective, be warm and supportive in tone, and incorporate the best practices from the research. Make it specific to the course subject if possible based on the syllabus content.]
    ```

    ---

    Please provide a thorough, evidence-based analysis that will help the instructor improve their syllabus to better support all students, especially those from historically marginalized groups."""

            # Call OpenAI API
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {
                        "role": "system",
                        "content": "You are an expert in educational equity and student-centered course design, specializing in analyzing syllabi for growth mindset and belonging cues based on research from the Student Experience Project."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                max_tokens=4000,
                temperature=0.7
            )

            # Extract the response text
            analysis_text = response.choices[0].message.content

            return {
                'analysis': analysis_text,
                'status': 'success'
            }

        except Exception as e:
            return {
                'error': f'Error calling OpenAI API: {str(e)}',
                'status': 'error'
            }

    def analyze_syllabus_quality(self) -> Dict[str, str]:
        """
        Use OpenAI GPT API to analyze syllabus for quality issues:
        - Undefined course terminology
        - Tone and inclusivity issues in policy language
        - Policies that may violate accessibility guidance or confuse students
        - Inconsistent formatting across repeated elements
        """
        # Extract full syllabus text with table structure preserved
        full_text = self._document_to_structured_text()

        # Get API key from environment
        api_key = os.environ.get('OPENAI_API_KEY')
        if not api_key:
            return {
                'error': 'OPENAI_API_KEY environment variable not set. Please configure your API key in .env file to enable quality analysis.',
                'status': 'error'
            }

        try:
            client = _get_openai()(api_key=api_key)

            prompt = f"""You are an expert in syllabus design, educational accessibility, and inclusive pedagogy. Analyze this course syllabus for quality and clarity issues across four key dimensions.

    SYLLABUS TEXT:
    {full_text}

    Please provide a comprehensive analysis covering these four areas:

    ## 1. UNDEFINED COURSE TERMINOLOGY

    Identify any technical terms, jargon, or course-specific vocabulary that students (especially first-time students in this field) might not understand. Look for:
    - Field-specific terminology used without definition
    - Acronyms that aren't spelled out on first use
    - Assignment types or formats that may be unfamiliar (e.g., "response paper", "think-pair-share", "flipped classroom")
    - Institutional terminology that may not be universally understood

    For each undefined term found, use this NESTED bullet format:
    - **Term**: [the term or phrase]
    - **Why it might be unclear**: [brief explanation]
    - **Suggested fix**: [how to define or clarify it]

    If no issues found, state: "✓ All course terminology appears to be clearly defined or would be familiar to students."

    ## 2. TONE AND INCLUSIVITY IN POLICY LANGUAGE

    Examine the tone of policies, rules, and requirements. Look for:
    - Punitive or threatening language (focus on consequences rather than support)
    - Deficit-based framing (assuming students will fail, cheat, or not try)
    - Language that may alienate students from marginalized groups
    - Overly legalistic or bureaucratic tone that creates distance
    - Policies that assume all students have the same resources, circumstances, or prior knowledge

    For each issue found, use this NESTED bullet format:
    - **Issue**: [quote the problematic language]
    - **Problem**: [why this language may be problematic]
    - **Better alternative**: [rewritten version with more inclusive, supportive tone]

    If no issues found, state: "✓ Policy language uses an inclusive, supportive tone throughout."

    ## 3. POLICIES THAT MAY CONFUSE OR EXCLUDE

    Identify policies or statements that:
    - Contradict accessibility best practices or may create barriers for students with disabilities
    - Conflict with institutional policies (if identifiable)
    - Are vague or ambiguous in ways that could confuse students
    - Make assumptions about student circumstances (e.g., access to technology, transportation, childcare)
    - May inadvertently exclude certain student populations

    For each policy issue, use this NESTED bullet format:
    - **Policy**: [quote the policy]
    - **Potential problem**: [what confusion or exclusion this might cause]
    - **Recommendation**: [how to revise or clarify]

    If no issues found, state: "✓ All policies appear clear and inclusive."

    ## 4. INCONSISTENT FORMATTING

    A FORMAT is the template or pattern used to express something — NOT the specific values filled in.
    A FORMAT INCONSISTENCY exists only when the TEMPLATE differs, not when the values differ.

    To test whether something is a format inconsistency, ask: "If I replaced the specific values with placeholders, would the two examples use a different template?" If no — it is NOT an inconsistency.

    CRITICAL RULE FOR DATES: Dates like "September 6", "September 13", and "September 28" all use the template [Month Name] [Day Number]. They are NOT inconsistent. The day number having 1 or 2 digits is NOT a format difference — numbers naturally vary in length. Only flag dates if the template itself differs, e.g., "September 6" vs "9/6" vs "Sept. 6th" (different templates).

    Examples of TRUE inconsistencies to flag:
    - Date FORMATS differ: "Jan 15" vs "1/15" vs "January 15th" (different templates: abbreviated-name vs numeric vs ordinal)
    - Headings at the same level using different styles (some bold, some underlined, different sizes)
    - Assignment names using inconsistent capitalization (some Title Case, some lowercase)
    - Times formatted differently: "2pm" vs "2:00 PM" vs "14:00"

    Examples that are NOT inconsistencies (do NOT flag these, and do NOT mention them at all):
    - "September 6" and "September 28" — same template [Month Day], different values
    - "September 6" and "September 13" — same template [Month Day], different values
    - "September 6" appearing in multiple places in the document — same template, same value
    - "Assignment 1" and "Assignment 2" — same template, different numbers
    - Multiple headings all using the same bold style

    If you are unsure whether something is a true format inconsistency, omit it entirely. Only include items that are unambiguously inconsistent.

    For each TRUE formatting inconsistency, use this NESTED bullet format:
    - **Element type**: [e.g., "dates", "headings", "assignment names"]
    - **Inconsistency**: [describe the actual template variation, showing contrasting examples]
    - **Recommendation**: [suggest a consistent format to use throughout]

    If no issues found, state: "✓ Formatting appears consistent throughout the syllabus."

    ---

    **IMPORTANT FORMATTING NOTE**:
    - Use NESTED bullets: main item as top-level bullet, sub-items indented with 4 spaces
    - Use clear markdown formatting with ## for main headings and **bold** for labels
    - Be specific and quote actual text from the syllabus
    - Prioritize the most significant issues in each category
    - If you find 5+ issues in a category, focus on the most important ones
    - DO NOT include a section on heading structure - that is covered separately in accessibility checks
    """

            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are an expert in syllabus design, educational accessibility, and inclusive pedagogy. Provide specific, actionable feedback."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=3000,
                temperature=0.2
            )

            analysis_text = response.choices[0].message.content

            return {
                'analysis': analysis_text,
                'status': 'success'
            }

        except Exception as e:
            return {
                'error': f'Error calling OpenAI API: {str(e)}',
                'status': 'error'
            }

    def analyze_image_alt_text(self) -> Dict[str, str]:
        """
        Use OpenAI GPT API to analyze image alt text quality:
        - Flag institutional logos without meaningful alt text
        - Identify alt text that is too short, irrelevant, or redundant (e.g., "Image of...")
        """
        # Extract all images with their alt text
        images_data = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph

            try:
                para_element = para._element

                # Look for modern drawing elements
                drawings = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')

                for drawing in drawings:
                    docPr = drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr')

                    if docPr is not None:
                        title = docPr.get('title', '')
                        descr = docPr.get('descr', '')
                        name = docPr.get('name', 'Unnamed image')

                        # Include images even if they have alt text (we want to check quality)
                        images_data.append({
                            'name': name,
                            'title': title,
                            'description': descr,
                            'location': para_info.location
                        })

                # Also check for older pict elements
                picts = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict')
                for pict in picts:
                    shape = pict.find('.//{urn:schemas-microsoft-com:vml}shape')
                    if shape is not None:
                        alt = shape.get('alt', '')
                        title = shape.get('title', '')

                        images_data.append({
                            'name': 'Legacy image',
                            'title': title,
                            'description': alt,
                            'location': para_info.location
                        })

            except:
                pass  # Skip if we can't parse images

        # If no images found, return early
        if not images_data:
            return {
                'analysis': 'No images found in the document.',
                'status': 'success'
            }

        # Get API key from environment
        api_key = os.environ.get('OPENAI_API_KEY')
        if not api_key:
            return {
                'error': 'OPENAI_API_KEY environment variable not set. Please configure your API key in .env file to enable image alt text analysis.',
                'status': 'error'
            }

        try:
            client = _get_openai()(api_key=api_key)

            # Format images data for the prompt
            images_list = []
            for idx, img in enumerate(images_data, 1):
                alt_text = img['description'] or img['title'] or '(no alt text)'
                images_list.append(f"{idx}. **{img['name']}** - Alt text: \"{alt_text}\" - Location: {img['location']}")

            images_text = "\n".join(images_list)

            prompt = f"""You are an expert in web accessibility and educational document design. Analyze the alt text quality for images in this syllabus.

    IMAGES AND THEIR ALT TEXT:
    {images_text}

    Please analyze these images for accessibility issues in two key areas:

    ## 1. INSTITUTIONAL LOGOS WITHOUT MEANINGFUL ALT TEXT

    Look for images that appear to be institutional logos, university seals, or organizational branding. These often have alt text like:
    - Empty or missing alt text
    - Just the filename (e.g., "logo.png")
    - Generic text like "logo" or "image"
    - The institution name without context (e.g., just "Harvard" instead of "Harvard University logo")

    **Why this matters**: Screen reader users need to know what the logo represents and its purpose in the document. Simply stating the institution name without context can be confusing.

    For each logo with inadequate alt text:
    - **Image**: [image name/number]
    - **Current alt text**: [what it currently says]
    - **Problem**: [why this is insufficient]
    - **Recommended alt text**: [specific suggestion, e.g., "University of Michigan logo" or "Michigan State University seal"]

    If no logo issues found, state: "✓ All institutional logos have meaningful alt text."

    ## 2. POOR QUALITY ALT TEXT

    Identify alt text that is:
    - **Too short**: Single words or very brief text that doesn't convey meaning (e.g., "chart", "graph", "picture")
    - **Redundant**: Starts with unnecessary phrases like "Image of...", "Picture of...", "Graphic showing..."
    - **Irrelevant**: Describes the image format rather than content (e.g., "JPEG image", "PNG file")
    - **Non-descriptive**: Vague text that doesn't convey the image's purpose (e.g., "illustration", "diagram")

    **Note**: Decorative images should have empty alt text, but this is rare in syllabi. Focus on informational images.

    For each poor quality alt text:
    - **Image**: [image name/number]
    - **Current alt text**: [what it currently says]
    - **Problem**: [specific issue - too short, redundant, irrelevant, or non-descriptive]
    - **Recommended alt text**: [improved version that clearly describes the image's content and purpose]

    If no issues found, state: "✓ All alt text is clear, concise, and descriptive."

    ---

    **IMPORTANT FORMATTING NOTES**:
    - Use clear markdown formatting with ## for main headings and **bold** for labels
    - Be specific and reference actual alt text from the images
    - Prioritize the most significant issues
    - If you find more than 5 issues in a category, focus on the most problematic ones
    - Remember: Good alt text is concise but informative, typically 1-2 sentences maximum
    """

            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are an expert in web accessibility, WCAG guidelines, and alt text best practices. Provide specific, actionable feedback."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=2000,
                temperature=0.7
            )

            analysis_text = response.choices[0].message.content

            return {
                'analysis': analysis_text,
                'status': 'success',
                'image_count': len(images_data)
            }

        except Exception as e:
            return {
                'error': f'Error calling OpenAI API: {str(e)}',
                'status': 'error'
            }

    def extract_course_metadata(self) -> Dict[str, str]:
        """
        Use OpenAI GPT API to extract course metadata from the syllabus.
        Returns instructor name, department, course subject, and course number.
        """
        # Extract first portion of syllabus text (metadata is usually at the top)
        syllabus_text = []
        for para_info in self.all_paragraphs[:50]:  # First 50 paragraphs should contain header info
            text = para_info.paragraph.text.strip()
            if text:
                syllabus_text.append(text)

        # Also check tables (course info is often in tables)
        for table in self.target_doc.tables[:3]:  # Check first 3 tables
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text.strip()
                    if text:
                        syllabus_text.append(text)

        full_text = "\n".join(syllabus_text)

        # Limit text length
        if len(full_text) > 4000:
            full_text = full_text[:4000]

        # Get API key from environment
        api_key = os.environ.get('OPENAI_API_KEY')
        if not api_key:
            return {
                'status': 'error',
                'error': 'OPENAI_API_KEY not set'
            }

        try:
            client = _get_openai()(api_key=api_key)

            prompt = f"""Extract the following information from this course syllabus. Return ONLY a JSON object with these exact keys. If a field cannot be determined, use null.

    SYLLABUS TEXT:
    {full_text}

    Extract:
    1. "instructor_name": The instructor's full name (e.g., "Dr. Jane Smith", "Professor John Doe", "Sarah Johnson, PhD")
    2. "department": The academic department (e.g., "Computer Science", "English", "Biology", "School of Business")
    3. "course_subject": The course subject code/prefix (e.g., "CS", "ENGL", "BIO", "MATH", "PSYCH")
    4. "course_number": The course number (e.g., "101", "350", "4400")
    5. "course_title": The full course title (e.g., "Introduction to Programming", "American Literature")
    6. "semester": The semester/term (e.g., "Fall 2024", "Spring 2025")

    Return ONLY valid JSON, no other text. Example:
    {{"instructor_name": "Dr. Jane Smith", "department": "Computer Science", "course_subject": "CS", "course_number": "101", "course_title": "Introduction to Programming", "semester": "Fall 2024"}}
    """

            response = client.chat.completions.create(
                model="gpt-4o-mini",  # Use faster/cheaper model for simple extraction
                messages=[
                    {"role": "system", "content": "You are a data extraction assistant. Extract course metadata from syllabi and return only valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=300,
                temperature=0.1  # Low temperature for consistent extraction
            )

            response_text = response.choices[0].message.content.strip()

            # Parse JSON response
            # Clean up response if needed (remove markdown code blocks)
            if response_text.startswith('```'):
                response_text = response_text.split('```')[1]
                if response_text.startswith('json'):
                    response_text = response_text[4:]
                response_text = response_text.strip()

            metadata = json.loads(response_text)
            metadata['status'] = 'success'
            return metadata

        except json.JSONDecodeError as e:
            return {
                'status': 'error',
                'error': f'Failed to parse metadata response: {str(e)}'
            }
        except Exception as e:
            return {
                'status': 'error',
                'error': f'Error extracting metadata: {str(e)}'
            }

