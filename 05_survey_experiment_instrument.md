# Survey Experiment Instrument

This file is designed so you can move it into Google Forms, Qualtrics, or another survey platform with minimal rewriting.

## 1. Survey title

**Student Evaluation of IT Micro-Credentials in Thailand**

## 2. Introductory text and consent

You are invited to participate in a research study about how university students in Thailand evaluate IT-related micro-credentials. This survey is for students currently enrolled in IT-related programs in Thailand. Participation is voluntary. Your responses will be anonymous and used only for academic research. You may stop at any time. The survey will take about 8 to 10 minutes.

By continuing, you confirm that:

- you are currently enrolled in a university in Thailand
- you are at least 18 years old
- you understand that your participation is voluntary

## 3. Screening questions

Use these first. Terminate the survey if the respondent is not eligible.

### SQ1

Are you currently enrolled in a university in Thailand?

- Yes
- No

### SQ2

What is your current field of study?

- Information Technology
- Computer Science
- Software Engineering
- Information Systems
- Computer Engineering
- Digital Business / Business Information Systems
- Other related digital or IT field
- Other non-IT field

Terminate if `Other non-IT field`.

### SQ3

Are you at least 18 years old?

- Yes
- No

Terminate if `No`.

## 4. Demographic and background questions

### D1

What is your current level of study?

- Undergraduate
- Postgraduate

### D2

What type of university do you currently attend?

- Public university
- Private university
- International university or international program
- Not sure

### D3

What year are you currently in?

- First year
- Second year
- Third year
- Fourth year
- Fifth year or above
- Postgraduate

### D4

Have you ever completed a MOOC or online short course with a certificate?

- Yes
- No
- Not sure

### D5

Have you ever completed a micro-credential, digital badge, or short professional certificate before?

- Yes
- No
- Not sure

### D6

Before today, how familiar were you with the term "micro-credential"?

- 1 = Not familiar at all
- 2 = Slightly familiar
- 3 = Moderately familiar
- 4 = Familiar
- 5 = Very familiar

## 5. Randomization instructions

After demographics, use a randomizer to assign each respondent evenly to one of the 8 vignette conditions below.

## 6. Base vignette setup

Keep this opening text constant for all conditions:

> Imagine that you are considering an online IT micro-credential called "Applied Data Analytics for Digital Projects."  
> The micro-credential is designed for university students and early-career learners who want to strengthen practical digital and IT skills.  
> It takes 20 hours to complete, includes online learning activities and an assessed final task, and provides a digital certificate after successful completion.

Then add the three experimental statements according to the assigned condition.

## 7. The 8 vignette conditions

### Condition 1

Recognition: No  
Stackability: No  
Industry endorsement: No

> This micro-credential is issued as an independent certificate and is not currently recognized by Thai universities or recorded in the National Credit Bank System.  
> It is a standalone certificate and cannot be combined with other short courses for future academic or professional credit.  
> It was developed by the provider alone and does not include endorsement from any industry association or employer group.

### Condition 2

Recognition: Yes  
Stackability: No  
Industry endorsement: No

> This micro-credential is formally recognized by participating Thai universities and can be recorded in the National Credit Bank System.  
> It is a standalone certificate and cannot be combined with other short courses for future academic or professional credit.  
> It was developed by the provider alone and does not include endorsement from any industry association or employer group.

### Condition 3

Recognition: No  
Stackability: Yes  
Industry endorsement: No

> This micro-credential is issued as an independent certificate and is not currently recognized by Thai universities or recorded in the National Credit Bank System.  
> It can be combined with other short courses and may later count toward a larger certificate or future academic credit pathway.  
> It was developed by the provider alone and does not include endorsement from any industry association or employer group.

### Condition 4

Recognition: No  
Stackability: No  
Industry endorsement: Yes

> This micro-credential is issued as an independent certificate and is not currently recognized by Thai universities or recorded in the National Credit Bank System.  
> It is a standalone certificate and cannot be combined with other short courses for future academic or professional credit.  
> It is endorsed by a recognized IT industry body and was reviewed with input from employers in digital and IT roles.

### Condition 5

Recognition: Yes  
Stackability: Yes  
Industry endorsement: No

> This micro-credential is formally recognized by participating Thai universities and can be recorded in the National Credit Bank System.  
> It can be combined with other short courses and may later count toward a larger certificate or future academic credit pathway.  
> It was developed by the provider alone and does not include endorsement from any industry association or employer group.

### Condition 6

Recognition: Yes  
Stackability: No  
Industry endorsement: Yes

> This micro-credential is formally recognized by participating Thai universities and can be recorded in the National Credit Bank System.  
> It is a standalone certificate and cannot be combined with other short courses for future academic or professional credit.  
> It is endorsed by a recognized IT industry body and was reviewed with input from employers in digital and IT roles.

### Condition 7

Recognition: No  
Stackability: Yes  
Industry endorsement: Yes

> This micro-credential is issued as an independent certificate and is not currently recognized by Thai universities or recorded in the National Credit Bank System.  
> It can be combined with other short courses and may later count toward a larger certificate or future academic credit pathway.  
> It is endorsed by a recognized IT industry body and was reviewed with input from employers in digital and IT roles.

### Condition 8

Recognition: Yes  
Stackability: Yes  
Industry endorsement: Yes

> This micro-credential is formally recognized by participating Thai universities and can be recorded in the National Credit Bank System.  
> It can be combined with other short courses and may later count toward a larger certificate or future academic credit pathway.  
> It is endorsed by a recognized IT industry body and was reviewed with input from employers in digital and IT roles.

## 8. Manipulation checks

Use a 5-point Likert scale from `1 = strongly disagree` to `5 = strongly agree`.

- MC1: This micro-credential appears to be formally recognized by institutions in Thailand.
- MC2: This micro-credential appears to be stackable toward a larger qualification or future credit.
- MC3: This micro-credential appears to have industry endorsement.

## 9. Student trust scale

Use a 5-point Likert scale from `1 = strongly disagree` to `5 = strongly agree`.

- TR1: I would trust this micro-credential as a legitimate learning qualification.
- TR2: This micro-credential appears credible to me.
- TR3: I feel confident that this micro-credential would represent real learning.
- TR4: I would regard this micro-credential as reliable evidence of skill development.
- TR5: This micro-credential appears worthy of serious consideration.

## 10. Perceived value scale

Use a 5-point Likert scale from `1 = strongly disagree` to `5 = strongly agree`.

- PV1: This micro-credential would be valuable for my academic or career development.
- PV2: This micro-credential would be worth the time required to complete it.
- PV3: This micro-credential would improve my learning or employability prospects.
- PV4: This micro-credential would be a useful addition to my qualifications.
- PV5: Overall, this micro-credential appears worthwhile.
- PV6: This micro-credential would provide meaningful benefits relative to the effort involved.

## 11. Enrollment intention scale

Use a 5-point Likert scale from `1 = strongly disagree` to `5 = strongly agree`.

- EI1: I would consider enrolling in this micro-credential if it were available to me.
- EI2: I would be interested in learning more about this micro-credential.
- EI3: I would be likely to pursue a similar micro-credential in the future.

## 12. Attention check

Include one item such as:

- AC1: To show that you are reading carefully, please select `Agree` for this statement.

## 13. Optional open-ended question

This is optional but can be useful for discussion and interpretation.

### O1

In one or two sentences, what most influenced your opinion of the micro-credential you just read about?

## 14. Suggested coding

### Experimental variables

- `REC`: 0 = no recognition, 1 = recognition
- `STACK`: 0 = not stackable, 1 = stackable
- `ENDORSE`: 0 = no industry endorsement, 1 = endorsement

### Composite variables

- `TRUST_MEAN`: average of TR1 to TR5
- `PVALUE_MEAN`: average of PV1 to PV6
- `EINTENT_MEAN`: average of EI1 to EI3

### Controls

- `PRIOR_MOOC`
- `PRIOR_MICROCRED`
- `STUDY_LEVEL`
- `UNI_TYPE`
- `MICROCRED_FAMILIARITY`

## 15. Pilot checklist

- Test whether students understand the term "stackable."
- Test whether recognition wording is clear in both English and Thai.
- Confirm that the vignette feels realistic and not overly promotional.
- Check average completion time.
- Revise any item that students find repetitive or confusing.

## 16. Recommended platform flow

### Google Forms

- Works for a pilot, but randomization is awkward.

### Qualtrics

- Best option for the full study because it supports randomization cleanly.

### Alternative

- Use 8 separate Google Form links if Qualtrics is unavailable, then distribute respondents evenly.

